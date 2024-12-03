"""
API Caller: Coded by @jjonamos
Authenticates using MWInit; makes
GET and POST requests to MAXIS API.
"""
import http.cookiejar as cookie_jar, json, os, pathlib, requests, tempfile, time, urllib3

##Disable warning about unverified HTTPS requests
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

class maxis:
    def __init__(self):
        self.__api_handler = requests.Session()
        self.__api_handler.verify = False
        self.__init_api_handler_cookies()

    def __init_api_handler_cookies(self, cookie_path = os.path.join(pathlib.Path.home(), '.midway', 'cookie')):
        #os.remove(cookie_path) ##For debugging purposes: Holy hand grenade existing cookies!
        if not os.path.exists(cookie_path):
            print('MIDWAY AUTHENTICATION\nInput will not display; press enter when done.')
            os.system('cmd /c "mwinit -s"') ##Run MWInit
        with tempfile.NamedTemporaryFile('w', delete = False) as temp_file:
            with open(cookie_path) as cookie_file:
                for line in cookie_file: ##Trim the 10 characters of '#HttpOnly_'
                    temp_file.write(line[10:] if line.startswith('#HttpOnly_') else line)
                    temp_file.flush()
                cookies = cookie_jar.MozillaCookieJar(temp_file.name)
                cookies.load()
        os.remove(temp_file.name)
        self.__api_handler.cookies = cookies
        self.__verify_cookies(cookie_path)
    
    def __verify_cookies(self, cookie_path):
        try:
            response = self.__api_handler.get('https://midway-auth.amazon.com/api/session-status')
            status = response.json()
            if not status['authenticated']: raise Exception
        except Exception:
            print('Retrying...')
            os.remove(cookie_path)
            self.__init_api_handler_cookies(cookie_path)

    def __call(self, call_type, call_function, max_attempts, backoff_factor):
        for attempt in range(max_attempts):
            try:
                response = call_function()
                response.raise_for_status()
                break ##Success; bravely run away!
            except Exception as e:
                num_attempt = attempt + 1
                sleep_time = num_attempt ** backoff_factor
                print(f'API {call_type} attempt {num_attempt}/{max_attempts} failed: {e}')
                if num_attempt < max_attempts: print(f'\nRetrying in {sleep_time} seconds...'); time.sleep(sleep_time)
                else: print('API call failed.')
        self.response = response.content.decode('utf-8')

    def get(self, api_path, api_authority = 'https://maxis-service-prod-pdx.amazon.com/', max_attempts = 10, backoff_factor = 2):
        call_function = lambda: self.__api_handler.get(api_authority + api_path)
        self.__call("GET", call_function, max_attempts, backoff_factor)

    def post(self, package, api_path, api_authority = 'https://maxis-service-prod-pdx.amazon.com/', max_attempts = 5, backoff_factor = 2):
        call_function = lambda: self.__api_handler.post(api_authority + api_path, json=json.loads(package), headers = {'Content-Type': 'application/json'})
        self.__call("POST", call_function, max_attempts, backoff_factor)