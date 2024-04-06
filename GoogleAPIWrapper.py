from apiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload
from consts import SCOPES
import os
import time


class GoogleAPIHandler:
    def __init__(self, input_target_folder_id):
        self.creds = None
        self.target_folder_id = input_target_folder_id
        self.slide_services = None
        self.drive_services = None
        self.sheet_services = None
        self.request_per_minute_counter = 0

    def _init_services(self):
        self.slide_services = build('slides', 'v1', credentials=self.creds)
        self.drive_services = build('drive', 'v3', credentials=self.creds)
        self.sheet_services = build('sheets', 'v4', credentials=self.creds)

    def init_api_handler(self):
        self.get_api_creds()
        self._init_services()

    # Creds
    def get_api_creds(self):
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.json'):
            self.creds = Credentials.from_authorized_user_file('token.json',
                                                               SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'secrets.json', SCOPES)
                self.creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(self.creds.to_json())

    # Slides
    def create_new_slideshow_from_template(self, slideshow_name, slideshow_id):
        file_metadata = {
            'name': slideshow_name,
            'parents': [self.target_folder_id]
        }

        new_slideshow = self.drive_services.files().copy(
            fileId=slideshow_id, body=file_metadata
        ).execute()

        new_slideshow_id = new_slideshow['id']

        return new_slideshow_id

    def get_presentation(self, presentation_id):
        # Call the Slides API
        presentation = self.slide_services.presentations().get(
            presentationId=presentation_id).execute()

        return presentation

    def get_slides(self, presentation_id):
        try:
            presentation = self.get_presentation(presentation_id)
            slides = presentation.get('slides')

            return slides

        except HttpError as err:
            print(err)
            return None

    def replace_text_in_presentation(self, text_to_replace, new_text, presentation_id):
        reqs = [
            {
                'replaceAllText':
                    {
                        'replaceText': new_text,
                        'containsText':
                            {
                                'text': text_to_replace,
                                'matchCase': True
                            }
                    }
            }
        ]
        return self.update_slideshow(reqs, presentation_id)

    def update_slideshow(self, reqs, presentation_id):
        if len(reqs) > 60:
            print('[GOOGLE API HANDLER LOG TEXT] --- Too many requests at once - splitting requests and trying again')
            half_point = int(len(reqs) / 2)
            half1_reqs = reqs[:half_point]
            half2_reqs = reqs[half_point:]
            self.update_slideshow(half1_reqs, presentation_id)
            self.update_slideshow(half2_reqs, presentation_id)
        can_update = self._add_request_to_counter(len(reqs))
        if can_update:
            response = self.slide_services.presentations().batchUpdate(body={'requests': reqs},
                                                                       presentationId=presentation_id,
                                                                       fields='').execute()
            return response

    @staticmethod
    def create_replace_image_request(image_id_to_replace, new_image_id):
        """
        :param image_id_to_replace: id of image in slides to replace
        :param new_image_id: Google id of image from URL
        """
        url_image_path = 'https://drive.google.com/uc?export=download&id=' + new_image_id

        request = {
            'replaceImage': {
                "imageObjectId": image_id_to_replace,
                "url": url_image_path
            }
        }

        return request

    # Drive
    def delete_file_from_drive(self, file_id):
        file = self.drive_services.files().delete(fileId=file_id).execute()

    # Sheets
    def read_sheet(self, sheet_id, cur_range=None):
        sheet = self.sheet_services.spreadsheets()
        result = sheet.values().get(spreadsheetId=sheet_id,
                                    range=cur_range).execute()
        values = result.get('values', [])

        return values

    # Generic Request Functions
    def _add_request_to_counter(self, number_of_requests_to_add):
        ready = False
        while not ready:
            if self.request_per_minute_counter + number_of_requests_to_add < 60:
                self.request_per_minute_counter += number_of_requests_to_add
                ready = True
            else:
                print('[GOOGLE API HANDLER LOG TEXT] --- exceeded 60 requests per minute limit waiting one minute')
                time.sleep(60)
                self.request_per_minute_counter = 0
                ready = False
        return ready
