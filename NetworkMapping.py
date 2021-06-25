#!/usr/bin/env python
# coding: utf-8

# In[20]:


### STOLEN FROM HOW TO USE GOOGLE SHEETS API https://developers.google.com/sheets/api/quickstart/python

import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

def main(spreadsheet_id):
    
    #RETRIEVES CREDENTIALS AND LOGS IN TO SHEET
    
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)

    
    MAIN_SPREADSHEET_ID = spreadsheet_id
    
    #SETS RANGES FOR THE DATA WE RETRIEVE
    RESPONSE_RANGE = 'Form Responses'
    PEOPLE_RANGE = 'Company Data!A2:B'
    QUESTION_RANGE = 'Company Data!C2:D'
    ROLE_RANGE = 'Company Data!F2:G'
    TEAM_RANGE = 'Company Data!I2:I'

    # Call the Sheets API
    sheet = service.spreadsheets()
    
    #RETRIEVE THE DATA ACCORDING TO THE RANGES ABOVE AND CONVERTS TO ARRAYS
    response = sheet.values().get(spreadsheetId=MAIN_SPREADSHEET_ID,range=RESPONSE_RANGE).execute()
    people = sheet.values().get(spreadsheetId=MAIN_SPREADSHEET_ID,range=PEOPLE_RANGE).execute()
    question = sheet.values().get(spreadsheetId=MAIN_SPREADSHEET_ID,range=QUESTION_RANGE).execute()
    role = sheet.values().get(spreadsheetId=MAIN_SPREADSHEET_ID,range=ROLE_RANGE).execute()
    team = sheet.values().get(spreadsheetId=MAIN_SPREADSHEET_ID,range=TEAM_RANGE).execute()
    
    response_values = response.get('values', [])
    people_values = people.get('values', [])
    question_values = question.get('values', [])
    role_values = role.get('values', [])
    team_values = team.get('values', [])
    
    return response_values, people_values, question_values, role_values, team_values


# In[21]:


#GET SPREADSHEET ID FROM spreadsheet_id.txt
with open('spreadsheet_id.txt') as f:
    spreadsheet_id = f.read()


# In[22]:


#RETRIEVE DATA FROM GOOGLE SHEET
response_values, people_values, question_values, role_values, team_values = main(spreadsheet_id)


# In[4]:


import pandas as pd
import numpy as np


# In[5]:


#CONVERTS ARRAYS FROM SHEETS API TO PANDAS DATAFRAMES
response_df = pd.DataFrame(response_values[1:], columns = response_values[0])
people_df = pd.DataFrame(people_values, columns = ["Index", "People"])
question_df = pd.DataFrame(question_values, columns = ["Question", "Role"])
role_df = pd.DataFrame(role_values, columns = ["Role", "Colour"])
team_df = pd.DataFrame(team_values, columns = ["Team"])


# In[6]:


#ALLOWS 3D DATAFRAMES WHERE PANDAS ONLY DOES 2D
import xarray as xr


# In[7]:


node_df = pd.DataFrame(0, columns = people_df["People"].values, index = team_df["Team"].values)

team_array = team_df["Team"].values
role_array = role_df["Role"].values
people_array = people_df["People"].values

#Each array is a person, each row is a team, each column is a role
edge_array = xr.DataArray(0, coords = [people_array, team_array, role_array], dims = ["People", "Team", "Role"])


team_question = question_df.loc[question_df["Role"] == "Team"].values[0][0]

for question in question_df["Question"].values:
    
    #DETERMINE THE ROLE ASSOCIATED WITH THE QUESTION
    role = question_df.loc[question_df["Question"] == question]["Role"].values[0]
    
    #CHECKING THAT ROLE IS NOT NONE TYPE (I.E FOR FACTUAL QUESTION (WHO IS BOSS?))
    if role in role_df["Role"].values:
        for person in people_df["People"].values:
            #IF PERSON IS AN ANSWER TO THE QUESTION
            if person in response_df[question].values:
                #RETRIEVE INDEX OF WHERE PERSON WAS THE ANSWER AND FIND THE TEAM WHO ANSWERED THE QUESTION
                index = response_df.index[response_df[question].values == person]
                describers = response_df.iloc[index][team_question].values
                #ADD A POINT TO THE PERSON FOR THE ROLE FOR THE TEAM THAT ANSWERED 
                #AND A POINT IN THE PERSON AND THE TEAM RELATION DATAFRAME
                for describer in describers:
                    edge_array.loc[(dict(People = person, Team = describer, Role = role))] = edge_array.loc[(dict(People = person, Team = describer, Role = role))] + 1
                    node_df[person][describer] = node_df[person][describer] + 1


# In[8]:


from pyvis.network import Network
import networkx as nx
import pydot


# In[9]:


network_map = nx.MultiDiGraph()


# In[10]:


#GENERATE GRAPH

#NODES ARE SIZED BY NUMBER OF LINKS TOTAL
#EDGES ARE SIZED BY NUMBER OF ANSWERS ASSOCIATED WITH EDGE

#PERSON -> ROLE -> TEAM

for i in range(0, len(people_df["People"].values)):
    if node_df.sum()[i].item() != 0:
        network_map.add_node(people_df["People"].values[i], size = node_df.sum()[i].item()*5, title = people_df["People"].values[i])

for i in range(0, len(team_df["Team"].values)):
    if node_df.sum(axis = 1)[i].item() != 0:
        network_map.add_node(team_df["Team"].values[i], size = node_df.sum(axis = 1)[i].item()*7, title = team_df["Team"].values[i], shape = "triangle", color=role_df.loc[role_df["Role"] == "Team"].values[0][1], physics = True)

for i in edge_array["People"].values:
    for j in edge_array["Role"].values:
        for k in edge_array["Team"].values:
            
            if edge_array.loc[(dict(People = i, Team = k, Role = j))].values.item() != 0:
                network_map.add_edge(
                    i, k, label = j, 
                    weight = edge_array.loc[(dict(People = i, Team = k, Role = j))].values.item(),
                    #VALUE = WIDTH
                    width = edge_array.loc[(dict(People = i, Team = k, Role = j))].values.item()*5,
                    title = j,
                    color = role_df.loc[role_df["Role"] == j].values[0][1]
                    )


# In[11]:


#GENERATE DOT FILE

nx.drawing.nx_pydot.write_dot(network_map, "network_map.dot")


# In[12]:


### DISPLAY GRAPH USING PYTHON GRAPHVIZ

rendered_map=Network(directed = True)
rendered_map.set_edge_smooth("dynamic")
rendered_map.from_nx(network_map)

neighbour_map = rendered_map.get_adj_list()
for node in rendered_map.nodes:
    if len(neighbour_map[node["id"]]) > 0:
        node["title"] += "<br>Linked to: " + "<br>".join(neighbour_map[node["id"]])


# In[13]:


rendered_map.set_options("""
var options = {
  "nodes": {
    "color": {
      "highlight": {
        "border": "rgba(222,233,96,1)"
      }
    },
    "font": {
      "size": 100,
      "background": "rgba(251,255,247,0)",
      "strokeWidth": 25
    }
  },
  "edges": {
    "color": {
      "highlight": "rgba(214,255,145,1)",
      "inherit": false
    },
    "font": {
      "size": 60,
      "background": "rgba(251,255,247,0)",
      "strokeWidth": 50
    },
    "smooth": {
      "forceDirection": "none"
    }
  },
  "physics": {
    "barnesHut": {
      "gravitationalConstant": -50000,
      "springLength": 4000,
      "springConstant": 0.0001,
      "avoidOverlap": 1
    },
    "minVelocity": 0.75
  }
}
""")


# In[14]:


rendered_map.save_graph("NetworkMap.html") 


# In[15]:


rendered_map.show("NetworkMap.html") 


# In[16]:


#CREATE DATAFRAME FOR RANKING

team_extend_df = np.array([])
role_extend_df = np.array([])
for i in range(0,len(role_df["Role"].values)-1):
    team_extend_df = np.append(team_extend_df, team_df["Team"].values)
team_extend_df.sort()
for i in range(0, len(team_df["Team"].values)):
    role_extend_df = np.append(role_extend_df, role_df["Role"].values[:-1])
rankings = pd.DataFrame({"Team":team} for team in team_extend_df)
rankings["Role"] = role_extend_df
for people in people_df["People"].values:
    rankings[people] = 0

for i in edge_array["People"].values:
    for j in edge_array["Role"].values:
        for k in edge_array["Team"].values:
            rankings.loc[((rankings["Team"].values == k) & (rankings["Role"].values == j)), i] = edge_array.loc[(dict(People = i, Team = k, Role = j))].values + rankings.loc[((rankings["Team"].values == k) & (rankings["Role"].values == j)), i]


# In[17]:


#SAVE STANDARD RANKINGS DATAFRAME
rankings.to_excel("rankings.xlsx")


# In[18]:


#TURN INTO PANDAS PIVOT TABLE
rankings_pivot = pd.pivot_table(rankings, index = ["Team", "Role"], aggfunc=np.sum).transpose()
rankings_pivot.to_excel("rankings_pivot.xlsx")


# In[19]:


#MAKE LEGEND FOR MAP
writer = pd.ExcelWriter("legend.xlsx", engine="xlsxwriter")
role_df.to_excel(writer, sheet_name="Legend")
                 
workbook = writer.book
worksheet = writer.sheets["Legend"]

for i, color in enumerate(role_df["Colour"].values):
    cell_format = workbook.add_format({
        "fg_color": color.lower()
    })
    worksheet.write(i+1, 2, color, cell_format)
writer.save()

