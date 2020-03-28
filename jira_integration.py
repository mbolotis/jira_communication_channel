from jira.client import JIRA
import pandas as pd


# jira_connection function initiates the connection to jira API
# the user should give valid credentials server, email, token
def jira_connection():
    server = str(input("Give me the name of the cloud server (without https// prefix) : "))
    server = "https://" + server
    mail = str(input("Give me your email : "))
    api_token = input("Give me your token : ")

    print("Input validation process started...")

    try:
        jira = JIRA(options={'server': server},
                basic_auth=(mail, api_token))
    except:
        jira = 0

    return jira


# data_extraction function takes as an argument the validated jira connection
# and it requests the JQL query by the user.
# Saves the outcome to a variable
def data_extraction(jira):
    jql = str(input("Provide the JQL query for the issues you want to extract : "))
    print("Please wait for the extraction to be completed")
    my_issues = jira.search_issues(jql)

    return my_issues


# Following the validation of the connection and the temp save of the JQL issues
# It initiates a Dataframe and appends all the requested issues in it
# It creates an .xlsx with the elements of the Dataframe
if __name__ == "__main__":
    connection = jira_connection()

    if connection == 0:
        print("Invalid input")

    else:
        issues = data_extraction(connection)

        fields = ['Key', 'Project', 'Assignee', 'Status', 'Created', 'Updated', 'Priority', 'Reporter', 'Hours_spent', 'Is_resolved']

        issues_df = pd.DataFrame(columns=fields)

        for issue in issues:
            if issue.fields.resolution is None:
                resolved = "No"
            else:
                resolved = "Yes"

            try:
                time_spent = issue.fields.aggregatetimespent / 3600
            except:
                time_spent = 0

            issues_df = issues_df.append({'Key': issue.key,
                                          'Project': issue.fields.project.name,
                                          'Assignee': issue.fields.assignee.displayName,
                                          'Status': issue.fields.status.name,
                                          'Created': issue.fields.created[0:10],
                                          'Updated': issue.fields.updated[0:10],
                                          'Priority': issue.fields.priority.name,
                                          'Reporter': issue.fields.reporter.displayName,
                                          'Hours_spent': time_spent,
                                          'Is_resolved': resolved
                                          }, ignore_index=True)

        issues_df.to_excel("jira_elements.xlsx", sheet_name="Jira Elements", index=False)

        print("Extraction process has been completed!")
