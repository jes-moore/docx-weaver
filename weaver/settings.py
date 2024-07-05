"""
Contains settings for DocxWeaver
"""

from typing import Literal
import logging
from pydantic import SecretStr
from pydantic_settings import BaseSettings

# Logging
log = logging.getLogger(__name__)

class DocxWeaverSettings(BaseSettings):
    """
    Settings class, loaded from environment
    """
    openai_api_key: SecretStr
    openai_model_name: Literal["gpt-4-turbo", "gpt-3.5-turbo", "gpt-4o"]
    def __init__(self, *a, **kw):
        super().__init__(*a, **kw)
        log.info("DocxWeaver Config: %s", self.model_dump_json())
