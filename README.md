# Two
Two Game

Two is a simulation of the card game (鋤大D) in the MS Excel environment. It is written in VBA. It features a multiplayer gameplay (ie. Players in different machines can interact with each other during the game).

Installation:
1.  Download Two.xlsm and Server.accdb from GitHub
2.  Put Server.accdb in a network location that all 4 players have read and write access. (eg. "\\pcXXX\folder\Server.accdb")
3.  Open Two.xlsm. Enable all macros. In worksheet [Setting], edit the location of access file (Cell C4) to the location of Server.accdb. (Change to "\\pcXXX\folder\Server.accdb" (no double quotes))
4.  Run macro "saveMultPlayer" to produce player 2 to 4 excel files (Two_2.xlsm, Two_3.xlsm, Two_4.xlsm) in the current folder.

Usage (steps can also be found in worksheet [Notes])
1.  Each player opens their excel file in their own machine.
2.  Player 1 checks the box "Restart", and then click the button "Resume".
3.  Other players then click the button "Resume".
4.  When it is your turn, make an action (play cards or pass).
5.  When the game ends, you can restart game from step 2.
