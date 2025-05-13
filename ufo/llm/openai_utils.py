import argparse
from mimetypes import guess_type

from msal import PublicClientApplication
import json
import requests
import base64
import time


class LLMClient:
    _ENDPOINT = 'https://fe-26.qas.bing.net/'
    _SCOPES = ['https://substrate.office.com/llmapi/LLMAPI.dev']
    _API = 'chat/completions'

    def __init__(self, endpoint):
        if endpoint != None:
            LLMClient._ENDPOINT = endpoint
        LLMClient._ENDPOINT += self._API

        self._app = PublicClientApplication('68df66a4-cad9-4bfd-872b-c6ddde00d6b2',
                                            authority='https://login.microsoftonline.com/72f988bf-86f1-41af-91ab-2d7cd011db47',
                                            enable_broker_on_windows=True, enable_broker_on_mac=True)

    def send_request(self, model_name, request):
        # get the token
        token = self._get_token()

        # populate the headers
        headers = {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + token,
            'X-ModelType': model_name}

        body = str.encode(json.dumps(request))
        response = requests.post(LLMClient._ENDPOINT, data=body, headers=headers)
        if (response.status_code != 200):
            raise Exception(f"Request failed with status code {response.status_code}. Response: {response.text}")
        return response.json()

    def send_stream_request(self, model_name, request):
        # get the token
        token = self._get_token()

        # populate the headers
        headers = {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + token,
            'X-ModelType': model_name}

        body = str.encode(json.dumps(request))
        response = requests.post(LLMClient._ENDPOINT, data=body, headers=headers, stream=True)
        for line in response.iter_lines():
            text = line.decode('utf-8')
            if text.startswith('data: '):
                text = text[6:]
                if text == '[DONE]':
                    break
                else:
                    yield json.loads(text)

    def _get_token(self):
        accounts = self._app.get_accounts()
        result = None

        if accounts:
            # Assuming the end user chose this one
            chosen = accounts[0]

            # Now let's try to find a token in cache for this account
            result = self._app.acquire_token_silent(LLMClient._SCOPES, account=chosen)

        if not result:
            result = self._app.acquire_token_interactive(scopes=LLMClient._SCOPES,
                                                         parent_window_handle=self._app.CONSOLE_WINDOW_HANDLE)

            if 'error' in result:
                raise ValueError(
                    f"Failed to acquire token. Error: {json.dumps(result, indent=4)}"
                )

        return result["access_token"]


# parser = argparse.ArgumentParser(description='Async API Example')
# parser.add_argument('--endpoint', type=str, help='Endpoint URL')
# parser.add_argument('--scenario', type=str, help='Scenario ID')
#
# args = parser.parse_args()
#
# endpoint = args.endpoint
# scenario_id = args.scenario

llm_client = LLMClient(None)


def send_request_img(model_name, image_path, prompt):
    encoded_image = base64.b64encode(open(image_path, 'rb').read()).decode('ascii')
    mime_type, _ = guess_type(image_path)
    if "o1" or "o3" in model_name:
        request_data = {
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": prompt
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:{mime_type};base64,{encoded_image}"
                            }
                        }
                    ]
                }
            ]
        }
    else:
        request_data = {
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": prompt
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:{mime_type};base64,{encoded_image}"
                            }
                        }
                    ]
                }
            ],
            "temperature": 0.01,
            "top_p": 0.95,
            "max_completion_tokens": 16384
        }

    # Available models are listed here: https://eng.ms/docs/experiences-devices/m365-core/microsoft-search-assistants-intelligence-msai/substrate-intelligence/llm-api/llm-api-partner-docs/available-models/available-models
    time_start = time.time()
    response = llm_client.send_request(model_name, request_data)
    result=response['choices'][0]['message']['content']
    time_end = time.time()
    # print(f"LLM Time taken: {time_end - time_start}")
    # print(result)
    return result


def send_request_ufo(model_name, message):
    request_data = {
            "messages":message ,
            "temperature": 0.01,
            "top_p": 0.95,
            "max_completion_tokens": 4096
    }

    # Available models are listed here: https://eng.ms/docs/experiences-devices/m365-core/microsoft-search-assistants-intelligence-msai/substrate-intelligence/llm-api/llm-api-partner-docs/available-models/available-models
    time_start = time.time()
    response = llm_client.send_request(model_name, request_data)
    result=response['choices'][0]['message']['content']
    time_end = time.time()
    # print(f"LLM Time taken: {time_end - time_start}")
    # print(result)
    return result

# prompt = "What’s in this image?"
# model_name = 'dev-gpt-45-preview'
# model_name = 'gpt-4o-vision-2024-05-13'
# image_path = r"C:\Users\v-yuhangxie\repos\20250506UIAgentUFO\ufo\llm\pic\try.jpg"
# send_request_img(model_name, image_path, prompt)
# # message=[]
# # send_request_ufo(model_name, message)