from pydantic import BaseModel


class LoginRequest(BaseModel):
    user_key: str = ""
    user_name: str = ""
    password: str


class UserSearchRequest(BaseModel):
    name: str
