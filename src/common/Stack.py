from typing import List, TypeVar, Generic

T = TypeVar('T')


class Stack(List[T], Generic[T]):

    def peek(self) -> T:
        if not self:
            raise IndexError("peek from an empty stack")
        return self[-1]
