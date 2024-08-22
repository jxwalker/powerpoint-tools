import logging
import asyncio
import json
import backoff
from asyncio_throttle import Throttler
from ibm_watson import NaturalLanguageUnderstandingV1
from ibm_watson.natural_language_understanding_v1 import Features, SummarizationOptions
from ibm_cloud_sdk_core.authenticators import IAMAuthenticator
from openai import OpenAI
from anthropic import Anthropic
from contextlib import asynccontextmanager
from config import Config

class SummarizationError(Exception):
    """Custom exception for summarization errors."""
    pass

def clean_summary(summary: str) -> str:
    lines = summary.split('\n')
    cleaned_lines = []

    for line in lines:
        # Strip whitespace from the beginning and end of the line
        stripped_line = line.strip()
        
        if stripped_line:
            # Remove all leading hyphens and spaces
            while stripped_line.startswith('-') or stripped_line.startswith(' '):
                stripped_line = stripped_line[1:].lstrip()
            
            # Add a single bullet point at the beginning
            stripped_line = f"- {stripped_line}"
            
            cleaned_lines.append(stripped_line)

    # Join the cleaned lines back into a single string
    cleaned_summary = "\n".join(cleaned_lines)
    
    # Replace any remaining double hyphens
    cleaned_summary = cleaned_summary.replace('- -', '-')
    
    return cleaned_summary

@asynccontextmanager
async def anthropic_client(config: Config):
    client = Client(api_key=config.anthropic_api_key)
    try:
        yield client
    finally:
        await client.aclose()

@asynccontextmanager
async def openai_client(config: Config):
    client = OpenAI(api_key=config.openai_api_key)
    try:
        yield client
    finally:
        await client.close()

@asynccontextmanager
async def watson_client(config: Config):
    try:
        authenticator = IAMAuthenticator(config.watson_api_key)
        natural_language_understanding = NaturalLanguageUnderstandingV1(
            version='2022-04-07',
            authenticator=authenticator
        )
        natural_language_understanding.set_service_url(config.watson_service_url)
        yield natural_language_understanding
    except Exception as e:
        logging.error(f"Failed to initialize Watson client: {str(e)}")
        raise SummarizationError(f"Watson client initialization failed: {str(e)}") from e

@backoff.on_exception(backoff.expo, Exception, max_tries=10)
async def summarize_with_claude(note: str, config: Config, summarization_level: int, throttler: Throttler) -> str:
    async with throttler:
        try:
            client = Anthropic(api_key=config.anthropic_api_key)
            prompt = f"Summarize the following notes into approximately {summarization_level} bullet points. Start each bullet point with '-' and ensure each is complete. Do not include any introductory text:\n\n{note}"
            
            response = await asyncio.to_thread(
                client.messages.create,
                model=config.anthropic_model,
                max_tokens=config.anthropic_max_tokens,
                messages=[
                    {"role": "user", "content": prompt}
                ]
            )
            
            if isinstance(response.content, list) and len(response.content) > 0:
                content = response.content[0]
                if hasattr(content, 'text'):
                    summary = content.text
                elif isinstance(content, dict) and 'text' in content:
                    summary = content['text']
                else:
                    summary = str(content)
            else:
                summary = str(response.content)
            
            return clean_summary(summary)
        except Exception as e:
            logging.error(f"Error in Claude summarization: {str(e)}")
            raise SummarizationError(f"Claude summarization failed: {str(e)}") from e

@backoff.on_exception(backoff.expo, Exception, max_tries=10)
async def summarize_with_openai(note: str, config: Config, summarization_level: int, throttler: Throttler) -> str:
    async with throttler:
        try:
            client = OpenAI(api_key=config.openai_api_key)
            response = await asyncio.to_thread(
                client.chat.completions.create,
                model=config.openai_model,
                messages=[
                    {"role": "system", "content": "You are a helpful assistant, skilled in summarizing complex documents into simple bullet points."},
                    {"role": "user", "content": f"Summarize the following notes into approximately {summarization_level} bullet points. Start each bullet point with '-' and ensure each is complete. Do not include any introductory text:\n\n{note}"}
                ],
                max_tokens=config.openai_max_tokens,
                temperature=config.openai_temperature
            )
            summary = response.choices[0].message.content.strip()
            return clean_summary(summary)
        except Exception as e:
            logging.error(f"Error in OpenAI summarization: {str(e)}")
            raise SummarizationError(f"OpenAI summarization failed: {str(e)}") from e

@backoff.on_exception(backoff.expo, Exception, max_tries=10)
async def summarize_with_watson(note: str, config: Config, summarization_level: int, throttler: Throttler) -> str:
    async with throttler:
        try:
            async with watson_client(config) as natural_language_understanding:
                response = await asyncio.to_thread(
                    natural_language_understanding.analyze,
                    text=note,
                    features=Features(summarization=SummarizationOptions(limit=summarization_level))
                )
                logging.debug(f"Watson API full response: {json.dumps(response.get_result(), indent=2)}")
                if 'summarization' in response.get_result() and 'text' in response.get_result()['summarization']:
                    summary = response.get_result()['summarization']['text']
                    logging.info(f"Watson summary: {summary}")
                    sentences = summary.split('.')
                    bullet_points = [f"- {sentence.strip()}" for sentence in sentences if sentence.strip()]
                    return "\n".join(bullet_points)
                else:
                    logging.warning(f"Watson API did not return a summary. Full response: {json.dumps(response.get_result(), indent=2)}")
                    raise SummarizationError("Watson API did not return a summary")
        except Exception as e:
            logging.error(f"Error in Watson summarization: {str(e)}")
            raise SummarizationError(f"Watson summarization failed: {str(e)}") from e