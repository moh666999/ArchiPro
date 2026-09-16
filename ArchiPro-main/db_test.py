import os
import libsql_client
from dotenv import load_dotenv

load_dotenv("ArchiPro.env") 

url = os.getenv("TURSO_DATABASE_URL")
auth_token = os.getenv("TURSO_AUTH_TOKEN")

client = libsql_client.create_client_sync(url=url, auth_token=auth_token)

result = client.execute("SELECT * FROM users")

for row in result.rows:
    print(row)