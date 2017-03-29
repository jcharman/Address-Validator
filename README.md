# Address Validator

Use getaddress.io to validate addresses stored in Google Forms.

# Usage

1) Create a service account in the Google Developer Console
2) Share Google Sheet with service account
3) Copy client secret to client_secret.json
4) Obtain api key from getaddress.io and place it in api_key.txt
5) Run the program as follows:

```python validate.py -r <sheet name>```
