CLIENT_ID = 'AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA'
CLIENT_SECRET = 'KKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKKK'

ENVIRONMENT = 'sandbox' #

REFRESH_TOKEN = 'TTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTTT'
COMPANY_ID = '0000000000000000'


try:
    from settings_local import *
except Exception as e:
    pass