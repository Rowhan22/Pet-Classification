# Author: Praveen Choudhary
# Email: praveen.choddhary1034@gmail.com
# Phone: 8619714798
# Organization: DisruptiveNext
# Supervisor: Mr. Prashant Mane


# Description: This code is part of an internship Task Attendance App. It interacts with the Microsoft Graph API
# to fetch calendar events using MSAL for authentication. The code uses the device code flow to obtain
# an access token and then fetches meetings from the user's calendar. The fetched meetings are filtered based
# on user-defined criteria such as subject, start date, and end date. The code displays key details of the events,
# including event ID, subject, start and end times, location, organizer information, and whether the event is an online meeting.
# Additionally, if the event is an online meeting, the code retrieves and displays the attendance report with participant information.

import msal
import requests
import asyncio
from datetime import datetime
from typing import Optional, List
from pydantic import BaseModel, Field, validator
import nest_asyncio

# Pydantic Models for input and output validation
class Attendance_App(BaseModel):
    """
    Pydantic model to define the filter criteria for calendar events.
    Allows filtering by subject, start date, and end date.
    """
    subject: Optional[str] = None
    start_date: Optional[datetime] = None
    end_date: Optional[datetime] = None

    @validator('start_date', pre=True)
    def validate_dates(cls, v):
        """
        Validates and converts the date string to a datetime object.
        """
        if isinstance(v, str):
            return datetime.strptime(v, "%Y-%m-%d")
        return v

class Organizer(BaseModel):
    """
    Pydantic model representing an event's organizer with their name and email.
    """
    name: str
    email: str

class EventDetails(BaseModel):
    """
    Pydantic model representing the details of an event including 
    the event ID, whether it's an online meeting, start and end times,
    location, organizer, and the online meeting URL.
    """
    id: str
    is_online_meeting: bool
    start_datetime: datetime
    end_datetime: datetime
    location_name: str
    organizer: Organizer
    join_url: str

class Event(BaseModel):
    """
    Pydantic model representing a basic event with an ID, subject, and whether it is an online meeting.
    """
    id: str
    subject: str
    isOnlineMeeting: bool

class MSALAuth:
    """
    Class to handle MSAL authentication for Azure AD, specifically using device code flow
    to obtain an access token for Microsoft Graph API.
    """
    def __init__(self, client_id: str, tenant_id: str):
        """
        Initializes the MSALAuth instance with the provided client ID and tenant ID.
        """
        self.client_id = client_id
        self.tenant_id = tenant_id
        self.authority = f"https://login.microsoftonline.com/{tenant_id}"

        # Initialize MSAL public client
        self.app = msal.PublicClientApplication(
            self.client_id,
            authority=self.authority
        )

    def get_access_token(self) -> str:
        """
        Authenticates the user and obtains an access token using device flow.
        
        Returns:
            str: Access token for Microsoft Graph API.

        Raises:
            Exception: If authentication fails.
        """
        flow = self.app.initiate_device_flow(scopes=["https://graph.microsoft.com/.default"])

        if "user_code" not in flow:
            raise Exception("Failed to create device flow. Exiting...")

        print(f"Please visit {flow['verification_uri']} and enter the code {flow['user_code']}.")

        result = self.app.acquire_token_by_device_flow(flow)

        if "access_token" in result:
            return result["access_token"]
        else:
            raise Exception(f"Failed to authenticate: {result.get('error_description')}")

async def get_events_with_filter(access_token: str, filter_data: Optional[Attendance_App] = None):
    """
    Fetches a list of calendar events from Microsoft Graph API, with optional filtering 
    by subject, start date, and end date.
    
    Args:
        access_token (str): The access token to authenticate the API request.
        filter_data (Optional[Attendance_App], optional): The filtering criteria for the events.

    Returns:
        dict: The response from the Graph API containing the events data.

    Raises:
        Exception: If the API request fails.
    """
    url = "https://graph.microsoft.com/v1.0/me/calendar/events"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    params = {
        "$select": "id,subject,isOnlineMeeting",
        "$top": "50"
    }

    if filter_data:
        filter_query = []
        if filter_data.subject:
            filter_query.append(f"subject eq '{filter_data.subject}'")
        if filter_data.start_date:
            filter_query.append(f"start/dateTime ge '{filter_data.start_date.isoformat()}'")
        if filter_data.end_date:
            filter_query.append(f"end/dateTime le '{filter_data.end_date.isoformat()}'")

        if filter_query:
            params["$filter"] = " and ".join(filter_query)

    response = requests.get(url, headers=headers, params=params)

    if response.status_code == 200:
        event_list = response.json()
        return event_list
    else:
        raise Exception(f"Failed to fetch events: {response.status_code}, {response.text}")

async def get_meeting_details(event_id: str, access_token: str) -> EventDetails:
    """
    Fetches detailed information about a specific event using its event ID.
    
    Args:
        event_id (str): The ID of the event to retrieve.
        access_token (str): The access token to authenticate the API request.

    Returns:
        EventDetails: A Pydantic model containing the event's details.
    
    Raises:
        Exception: If the API request fails.
    """
    url = f"https://graph.microsoft.com/v1.0/me/events/{event_id}"

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        event = response.json()

        organizer = Organizer(
            name=event['organizer']['emailAddress']['name'],
            email=event['organizer']['emailAddress']['address']
        )

        event_details = EventDetails(
            id=event['id'],
            is_online_meeting=event.get('isOnlineMeeting', False),
            start_datetime=datetime.fromisoformat(event['start']['dateTime']),
            end_datetime=datetime.fromisoformat(event['end']['dateTime']),
            location_name=event.get('location', {}).get('displayName', 'No location specified'),
            organizer=organizer,
            join_url=event.get('onlineMeeting', {}).get('joinUrl', 'No online meeting URL')
        )
        return event_details
    else:
        raise Exception(f"Failed to fetch event details: {response.status_code}, {response.text}")

def display_event_details(event_details: EventDetails):
    """
    Displays the details of a meeting, including its ID, start time, end time, location, 
    organizer, and online meeting URL.
    
    Args:
        event_details (EventDetails): The details of the event to display.
    """
    print("-" * 50)
    print("Meeting Details/Report:")
    print("-" * 50)
    print(f"ID: {event_details.id}")
    print(f"Is Online Meeting: {event_details.is_online_meeting}")
    print(f"Start: {event_details.start_datetime}")
    print(f"End: {event_details.end_datetime}")
    print(f"Location: {event_details.location_name}")
    print(f"Organizer: {event_details.organizer.name} ({event_details.organizer.email})")
    print(f"Join URL: {event_details.join_url}")
    print("-" * 50)

def get_attendance_report_from_url(join_url, access_token):
    """
    Fetches the attendance report for an online meeting using its join URL.
    
    Args:
        join_url (str): The join URL of the online meeting.
        access_token (str): The access token to authenticate the API request.
    """
    url = f"https://graph.microsoft.com/v1.0/me/onlineMeetings?$filter=joinWebUrl eq '{join_url}'"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        meetings = response.json().get("value", [])
        if meetings:
            for meeting in meetings:
                meeting_id = meeting.get("id")
                print(f"Meeting ID: {meeting_id}")
                get_attendance_report(meeting_id, access_token)
        else:
            print("No meetings found with the specified Join URL.")
    else:
        print(f"Error {response.status_code}: {response.text}")

def get_attendance_report(meeting_id, access_token):
    """
    Fetches the attendance report for a specific online meeting using its meeting ID.
    
    Args:
        meeting_id (str): The ID of the meeting to get the attendance report for.
        access_token (str): The access token to authenticate the API request.
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    url = f"https://graph.microsoft.com/v1.0/me/onlineMeetings/{meeting_id}/attendanceReports"
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        attendance_data = response.json()
        print("-" * 50)
        print("Attendance Report:")
        print("-" * 50)
        for report in attendance_data.get("value", []):
            meeting_start = report.get("meetingStartDateTime")
            meeting_end = report.get("meetingEndDateTime")
            participant_count = report.get("totalParticipantCount")
            print(f"Attendance ID: {report.get('id')}")
            print(f"Start Time: {meeting_start}")
            print(f"End Time: {meeting_end}")
            print(f"Total Participants: {participant_count}")
            print("-" * 50)
    else:
        print(f"Error {response.status_code}: {response.text}")

async def print_events(CLIENT_ID: str, TENANT_ID: str, filter_data: Optional[Attendance_App] = None):
    """
    Fetches and displays calendar events, then fetches and displays meeting details
    and attendance reports for a specific event.
    
    Args:
        CLIENT_ID (str): The client ID for authentication.
        TENANT_ID (str): The tenant ID for authentication.
        filter_data (Optional[Attendance_App], optional): The filtering criteria for events.
    """
    # Initialize MSAL Authentication
    auth = MSALAuth(CLIENT_ID, TENANT_ID)
    access_token = auth.get_access_token()

    # Fetch events with or without filter
    event_list = await get_events_with_filter(access_token, filter_data)

    if event_list.get("value"):
        events = [Event(id=event['id'], subject=event['subject'], isOnlineMeeting=event['isOnlineMeeting']) for event in event_list["value"]]
        for event in events:
            print(f"Event ID: {event.id}\nSubject: {event.subject}\nIs Online Meeting: {event.isOnlineMeeting}")
    else:
        print("No events found.")

    # Asking for event details
    event_id = input("Enter the Event ID to get meeting details: ")
    event_details = await get_meeting_details(event_id, access_token)
    display_event_details(event_details)

    # Now, proceed to second code for attendance report
    join_url = event_details.join_url
    if join_url != 'No online meeting URL':
        print("\nFetching attendance report...\n")
        get_attendance_report_from_url(join_url, access_token)
    else:
        print("No online meeting URL available to fetch attendance report.")

# Run the asyncio event loop
if __name__ == "__main__":
    nest_asyncio.apply()  # To enable nested event loops if running in Jupyter or other environments

    # Define your credentials here
    CLIENT_ID = "1c8a96b8-2a14-4305-9589-1ef3cece0729"
    TENANT_ID = "32600c8e-9edb-4773-92ff-bd4586b0691f"

    # Ask the user if they want to apply a filter
    filter_response = input("Do you want to apply a filter (yes/no)? ").strip().lower()

    filter_data = None
    if filter_response == 'yes':
        subject = input("Enter the subject to filter (leave blank for no filter): ").strip()
        start_date_str = input("Enter the start date (YYYY-MM-DD, leave blank for no filter): ").strip()
        end_date_str = input("Enter the end date (YYYY-MM-DD, leave blank for no filter): ").strip()

        start_date = datetime.strptime(start_date_str, "%Y-%m-%d") if start_date_str else None
        end_date = datetime.strptime(end_date_str, "%Y-%m-%d") if end_date_str else None

        filter_data = Attendance_App(subject=subject, start_date=start_date, end_date=end_date)

    # Run the event fetching process
    asyncio.run(print_events(CLIENT_ID, TENANT_ID, filter_data))
