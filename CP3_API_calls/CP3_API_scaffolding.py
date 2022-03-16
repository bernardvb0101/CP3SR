from CP3_API_calls.CP3_API_Classes import CP3Client

def get_token(username, password, grant_type, full_url):
    client = CP3Client(username, password, grant_type, full_url)
    token = client.get_API_token
    return token

