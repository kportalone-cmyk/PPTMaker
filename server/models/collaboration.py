from pydantic import BaseModel
from typing import Optional


class CollaboratorAdd(BaseModel):
    user_key: str
    role: str = "editor"  # "editor" | "viewer"


class CollaboratorUpdate(BaseModel):
    role: str  # "editor" | "viewer"
