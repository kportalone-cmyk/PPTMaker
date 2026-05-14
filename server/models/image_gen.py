"""이미지 생성 워크스페이스(image_gen project_type) 요청 모델"""

from pydantic import BaseModel, Field
from typing import Literal, Optional


class ImageGenRequest(BaseModel):
    """POST /api/image-gen/{project_id}/generate 요청 본문"""
    prompt: str
    model: Literal["nano-banana-pro", "nano-banana-2", "imagen-4"]
    aspect_ratio: Literal["16:9", "4:3", "1:1", "3:4", "9:16"]
    count: int = Field(ge=1, le=4)
    # @멘션으로 선택된 참조 이미지의 file_id 들. 동일 프로젝트의 image_generations 안에서 조회.
    # nano-banana 계열만 지원, imagen-4 와 함께 보내면 400 응답.
    reference_file_ids: Optional[list[str]] = None
