import requests
from requests.auth import HTTPBasicAuth

PUBLIC_KEY = "53445caa91"
PRIVATE_KEY = "5e75212c8ae9d7e"

url = "https://watson.is/gravityformsapi/forms"

r = requests.get(
    url,
    auth=HTTPBasicAuth(PUBLIC_KEY, PRIVATE_KEY)
)

print(r.status_code)
print(r.text[:1000])