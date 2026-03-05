from pydantic import BaseModel


class LoginRequest(BaseModel):
    user_key: str
    password: str


class UserSearchRequest(BaseModel):
    name: str
