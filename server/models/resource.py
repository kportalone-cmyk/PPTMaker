from pydantic import BaseModel
from typing import Optional


class ResourceCreate(BaseModel):
    project_id: str
    resource_type: str  # "file" | "text" | "web"
    title: str = ""
    content: Optional[str] = None
    file_path: Optional[str] = None
    source_url: Optional[str] = None


class WebSearchRequest(BaseModel):
    project_id: str
    query: str


class URLResourceRequest(BaseModel):
    project_id: str
    urls: list
