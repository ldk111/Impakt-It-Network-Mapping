Instructions For Use:

1. From the ImpaktIt Question Sheet spreadsheet make sure that the form is linked to the Form Responses page so that it auto updates.

2. In the form I have included a script that will automatically populate the allowed answers for a question.
	
	On the "Questionnaire Questions" sheet you see that it has all the questions in the form along the top as columns and the possible responses as rows beneath each question.
	These questions are retrieved from the "Company Data" sheet under the "Questions" column, they will automatically be added to the head of the columns in the "Questionnaire Questions" sheet.
		The questions in that list must be the EXACT SAME as the questions on the form and it must be in the SAME ORDER.
		The questions themselves do not get inputted to the form, only the answers, questions must be added manually to the form and the sheet.
	To access the script go to Tools -> Script Editor, then run the populateGoogleForms function to populate the answer fields of the form.

3. In the "Company Data" sheet the following columns must remain in the same place and only expand downwards;
	Index, People, Question, Role, Role, Colour, Team
	
	Index is just the number of the person, People is the list of names that are in the company. These need to be updated for each company.
	Question and the first Role are the questions asked in the form and the associated role they are inferring i.e if the question was asking "Who would be a good agent?" the Role would be "Agent"
		If Role is blank the question will not be used in generating the map i.e for questions where the response is not just a single name from the People list.
		The question asking which team they are from MUST have the role "Team" and the responses must be only elements of the "Team" column as this question will define which team the links people have are between.
	The second Role and Colour define the possible roles that people can take in the company e.g "Agent" or "Sponsor". Colour will define the colour of the edge that represents this role.
	For example, if the Role said "Sponsor" and the Colour was "Blue", all the lines that represented a "Sponsor" link in the network map would be coloured "Blue".
		Note only use simple basic colours, a full list of supported colours can be found at https://xlsxwriter.readthedocs.io/working_with_colors.html?highlight=color
	Team is the available teams in the company that a person can say they are from. 
		These will be nodes along with People and will function as the user identifier to define where a link is from.

4. Once questions are populated in the google form it can be sent out to users for them to fill in.

5. Before you run the network mapping script you must first fill in the spreadsheet_id.txt file with the id of the ImpaktIt Question Sheet google sheet. 
	i.e if the link to the sheet is https://docs.google.com/spreadsheets/d/1M8wjJekcNPdQ7bW1UsU_LCvXbIfo8qItQNrJjU1WoGg/edit#gid=679243854 the id is 1M8wjJekcNPdQ7bW1UsU_LCvXbIfo8qItQNrJjU1WoGg
	and that string needs to go in the spreadsheet_id.txt file and that string alone.

6. You will also need credentials in order to access the sheet with the code. These can be gotten from google cloud platform under APIs and Services under Credentials and you need to download the OAuth 2.0 Client IDs
	Then you need to rename this .json file to credentials.json and place in the folder of the python file.
	I used the google sheets api.
	Further information on how I did this is under https://developers.google.com/sheets/api/quickstart/python, https://developers.google.com/workspace/guides/create-credentials, we need desktop credentials so this subsection https://developers.google.com/workspace/guides/create-credentials#desktop is what I am referring to.

7. Once the credentials and spreadsheet ID are in the same folder you should be able to run the python file providing you have the required dependencies and it will output:
	A NetworkMap.html file for a rendering of the map.
	A network_map.dot containg the dot information of the map for rendering elsewhere.
	A rankings.xlsx containing how many links each person got for each role in each team.
	A rankings_pivot.xlsx containing the same data as above in a different format.
	A legend.xlsx containing the legend for the edges of the network map.

Code Dependencies:

I was running this using a conda environment of Python 3.8.5 and the packages explicitly used in the code:

google-api-python-client
google-auth-httplib2
google-auth-oauthlib
pandas
numpy
xarray
pyvis
networkx
pydot


