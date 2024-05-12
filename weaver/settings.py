"""
Contains settings for WordWeaver
"""

from typing import Literal
import logging
from pydantic import SecretStr
from pydantic_settings import BaseSettings

# Logging
log = logging.getLogger(__name__)

class WordWeaverSettings(BaseSettings):
    """
    Settings class, loaded from environment
    """
    openai_api_key: SecretStr
    openai_model_name: Literal["gpt-4-turbo", "gpt-3.5-turbo"] = "gpt-4-turbo"
    def __init__(self, *a, **kw):
        super().__init__(*a, **kw)
        log.info("WordWeaver Config: %s", self.model_dump_json())
