# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.


from .basic import BaseSession
from .state import (ErrorState, MaxStepReachedState, NoneState,
                           SessionFinishState)
from typing import List


class UFOClient:

    def __init__(self, session: BaseSession) -> None:
        """
        Initialize a UFO client.
        """

        self.session = session


    def run(self) -> None:
        """
        Run the UFO client.
        """

        while not self.session.is_finish():
            self.session.handle()

        # If the session is finished normally, handle the last state with additional logic.
        if isinstance(self.session.get_state(), SessionFinishState):
            self.session.handle()
    
        self.session.print_cost()



class UFOClientManager:

    def __init__(self, session_list: List[UFOClient]) -> None:
        """
        Initialize a batch UFO client.
        """

        self.session_list = session_list


    def run_all(self) -> None:
        """
        Run the batch UFO client.
        """

        for session in self.session_list:
           client = UFOClient(session)
           client.run()