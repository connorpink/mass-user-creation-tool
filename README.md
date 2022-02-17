# mass-user-creation-tool
Command-line tool using csv import to mass-generate net new Active Directory users

In this repo is the csv file you can use as a template.
CSV should look like this: 
| firstname       | lastname     | email             | displayname  | username  |
| :-----------    | :----------- | :--------------   | :---------   | :---------|
| john            | doe          | jdoe@email.com    | doe, john    | jdoe      |
| tom             | smith        | tsmith@email.com  | smith, tom   | tsmith    |


- If username is already taken in directory using format '(firstname[0])(lastname)' program will create user with more letters from first name list this: 
'(firstname[0-1])(lastname)'.... '(firstname[0-2])(lastname)'.... '(firstname[0-3])(lastname)' until username is available. 
  - If firstname runs out of characters to add the program will throw an error and skip that user so that it can be made manually.
- Program generates a Log file in the directory location of the script that logs all user creations info with timestamps. If user was unable to be created it is listed in log as well.
- Users that have a last name with hyphons such as 'sung-lee' can be left alone in the spreadsheet and the program will handle them with the following format for usernames:
'sung-lee' --> 'slee', and 'sung-jung-lee' --> 'sjlee'.
