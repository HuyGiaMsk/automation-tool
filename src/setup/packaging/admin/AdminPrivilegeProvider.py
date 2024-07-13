from abc import ABC, abstractmethod


class AdminPrivilegeProvider(ABC):

    @abstractmethod
    def is_currently_running_with_admin(self) -> bool:
        pass

    @abstractmethod
    def provide_admin_privilege(self) -> None:
        pass
