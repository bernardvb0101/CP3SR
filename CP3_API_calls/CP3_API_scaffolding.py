from CP3_API_calls.CP3_API_Classes import CP3Client

def get_token(username, password, grant_type, url_choice):
    client = CP3Client(username, password, grant_type, url_choice)
    token = client.get_API_token
    return token

