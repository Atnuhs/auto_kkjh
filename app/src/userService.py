from typing import Union
from src.user import User
from src.userRepository import UserRepository


class UserSettingService:
    def isDpulicated(self, user: User) -> Union[None, User]:
        userRepository = UserRepository()
        return userRepository.find(user.userName)
