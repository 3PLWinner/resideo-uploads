import os
import io
import requests
import pandas as pd
import csv
from io import StringIO
import base64
from dotenv import load_dotenv
from msal import ConfidentialClientApplication
import json
import datetime

class Email:
    email_id = None
    # Email JSON needed for Microsoft Graph API

    email_json =  {
        "subject" : "Outbound Resideo Orders",
        "body":{
            "contentType" : "HTML",
        },
        "toRecipients" : [
            {
                "emailAddress" : {
                    "address" : "williamson.alex@3plwinner.com"
                }
            },
            {
                "emailAddress" : {
                    "address" : "limbach.june@3plwinner.com"
                }
            }
        ]
    }

    # Empty method
    def generate_email():
        return ""

# Child class specifically for order email
class OrderEmail(Email):
    # Dictionary that has date for key and an array of order IDs for the email
    date_dict = {}

    # Adds new order ids to the date_dict
    def add_to_body(self, orders):
        for order in orders:
            # If it is in the dictionary
            if not(self.date_dict.get(order[-1]) is None):
                # If the order ID is not already in the array
                if not(order[0] in self.date_dict[order[-1]]):
                    # Add the order ID to the date
                    self.date_dict[order[-1]].append(order[0])
            else:
                # Create new date key and add the order id array
                self.date_dict[order[-1]] = [order[0]]

    # Generates the order email
    def generate_email(self):

        body_html = ""

        # Creates a list of dates to be sorted
        dates = list(self.date_dict.keys())
        dates.sort()

        # Takes sorted dates and adds the html lines for the orders
        for date in dates:
            order_ids = self.date_dict[date]

            body_html += f"<p><u>{date}</u><p>"

            for order_id in order_ids:
                body_html += f"<p>{order_id}</p>"
        
        # Add the new body html to the email_json
        self.email_json['body']['content'] = body_html
        
        # Return prepared dict object
        return self.email_json
    
    def has_orders(self):

        dates = self.date_dict.keys()

        if len(dates) == 0:
            return False

        return True

# Child of Email class for Error Email
class ErrorEmail(Email):
    # A dictionary that uses order ID for keys and error text for value
    error_dict = {}
    hasError = False

    def __init__(self):
        self.offers = []

        # Sends this to the Resideo email
        self.email_json["toRecipients"] = [
            {
                "emailAddress" : {
                    "address" : "resideo@3plwinner.com"
                }
            }
        ]

    # Add error message to the error_dict under the order id
    def add_to_body(self, order_id, error_message):

        if not(self.error_dict.get(order_id) is None):
            self.error_dict[order_id] += "\n\n"
            self.error_dict[order_id] += error_message
        else:
            self.error_dict[order_id] = error_message

    def generate_email(self):

        date_string = datetime.datetime.now().strftime("%m-%d-%Y at %H:%M")

        # Adds new subject to error email
        self.email_json["subject"] = f"Resideo Order Errors {date_string}"

        body_html = ""

        # Iterates through the order ids and creates the html for an error code
        for order in self.error_dict.keys():
            errors = self.error_dict[order]

            body_html += f"<p><u>The order with ID {order}</u><p>"
            body_html += f"<p>Had the following errors:<p>"
            body_html += f"<p>{errors}</p>"
            body_html += "<br>"
        
        # Adds to the body of the email JSON
        self.email_json['body']['content'] = body_html
        
        return self.email_json

    # Adds offers 
    def add_offers(self, offers):     
        for offer in offers:
            self.offers.append(offer)
    
    # Generates bytes for CSV attachment
    def generate_error_bytes(self):
        # Inserts the headers as the first tuple
        self.offers.insert(0,('Delivery Number', 'Company Name/Contact Name', 'Address 1', 'Address 2', 'Address 3',
             'City', 'State', 'Postal Code', 'Country', 'Product ID', 'Quantity', 'Sales Order',
             'Shipping Conditions', 'Delivery Instructions', 'Carrier', 'Planned Ship Date'))

        # Creates an IO object and adds it to the CSV writer
        attachment_string = StringIO()

        csv_writer = csv.writer(attachment_string)

        # Writes each tuple in csv format
        for offer in self.offers:
            csv_writer.writerow(offer)

        csv_string = attachment_string.getvalue()

        # Encodes to base64 and gets the string to add to the JSON body
        encoded_csv = base64.b64encode(csv_string.encode("utf-8"))
        encoded_string = encoded_csv.decode("utf-8")

        attachment_string.close()

        return encoded_string
        
# Orders class to generate XML API calls to VeraCore
class Orders:
    offers = []
    purchase_orders = []
    
    def __init__(self, order_id= None):
        self.order_id = order_id
        self.offers = []

    def add_to_offers(self, offer):
        self.offers.append(offer)

    # Iterates through added offers and creates the offer XML to be added
    def private_generate_offer_xml(self):

        offer_string = ""
        purchase_order_string = ""

        for index, offer in enumerate(self.offers):
            new_offer = f"""
                    <OfferOrdered>
                        <Offer>
                            <Header>
                                <ID>{offer[9]}</ID>
                            </Header>
                        </Offer>
                        <Quantity>{int(offer[10])}</Quantity>
                        <OrderShipTo>
                            <Key>1</Key>
                        </OrderShipTo>
                    </OfferOrdered>"""
            offer_string += new_offer

            # Adds all the purchase order numbers to one string
            if not(offer[11] in self.purchase_orders):
                
                if index == len(self.offers)-1:
                    purchase_order_string += str(offer[11])
                else:
                    purchase_order_string += str(offer[11]) + ","
                
                self.purchase_orders.append(offer[11])
        

        return offer_string, purchase_order_string
    
    # Generates XML needed for VeraCore SOAP API Add Orders endpoint
    def generate_xml(self):
        offer_string, purchase_order_string = self.private_generate_offer_xml()

        return f"""<?xml version="1.0" encoding="utf-8"?>
        <soap:Envelope
            xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
            xmlns:xsd="http://www.w3.org/2001/XMLSchema">
            <soap:Header>
                <AuthenticationHeader
                    xmlns="http://omscom/">
                    <Username>VSO335</Username>
                    <Password>testapiuser123</Password>
                </AuthenticationHeader>
            </soap:Header>
            <soap:Body>
                <AddOrder
                    xmlns="http://omscom/">
                    <order>
                        <Header>
                            <ID>{self.order_id}</ID>
                            <EntryDate>2025-07-16T00:00:00</EntryDate>
                            <Comments>{self.offers[0][13]}</Comments>
                            <ReferenceNumber>{purchase_order_string}</ReferenceNumber>
                        </Header>
                        <Shipping>
                            <FreightCarrier>
                                <Name>{self.offers[0][12]}</Name>
                            </FreightCarrier>
                            <NeededBy>{self.offers[0][15]}</NeededBy>
                        </Shipping>
                        <Money></Money>
                        <Payment></Payment>
                        <OrderedBy>
                            <CompanyName>{self.offers[0][1]}</CompanyName>
                            <Address1>{self.offers[0][2]}</Address1>
                            <Address2>{self.offers[0][3]}</Address2>
                            <Address3>{self.offers[0][4]}</Address3>
                            <City>{self.offers[0][5]}</City>
                            <State>{self.offers[0][6]}</State>
                            <PostalCode>{self.offers[0][7]}</PostalCode>
                            <Country>{self.offers[0][8]}</Country>
                        </OrderedBy>
                        <ShipTo>
                            <OrderShipTo>
                                <Flag>OrderedBy</Flag>
                                <Key>1</Key>
                            </OrderShipTo>
                        </ShipTo>
                        <BillTo>
                            <Flag>OrderedBy</Flag>
                        </BillTo>
                        <Offers>
                        {offer_string} 
                        </Offers>
                    </order>
                </AddOrder>
            </soap:Body>
        </soap:Envelope>
        """

# Writes Microsoft errors to error log
def write_to_log(text):
    path = os.getcwd()
    with open(path+"/"+"errors.txt", "a") as file:
        file.write(datetime.datetime.now().strftime("--------%m-%d-%yT%H:%M:%S----------------\n\n"))
        file.write(text)

# Processes raw df that is included in the attachment
def process_df(df):
    
    # Convert Planned Shipped Date from yyyymmdd to mm/dd/yyyy
    df['Planned Ship Date'] = pd.to_datetime(df['Planned Ship Date'], format='%Y%m%d').dt.strftime('%m/%d/%Y')
    
    # Move the Carrier column data to Shipping Conditions column
    df['Shipping Conditions'] = df['Carrier']
    
    # Clear the original Carrier column but keep it in the DataFrame
    df['Carrier'] = ''
    
    # Group by Delivery Number, Product ID, and aggregate the Quantity
    df = df.groupby(['Delivery Number', 'Product ID'], as_index=False).agg({
        'Company Name/Contact Name': 'first',
        'Address 1': 'first',
        'Address 2': 'first',
        'Address 3': 'first',
        'City': 'first',
        'State': 'first',
        'Postal Code': 'first',
        'Country': 'first',
        'Quantity': 'sum',
        'Sales Order': 'first',
        'Shipping Conditions': 'first',
        'Delivery Instructions': 'first',
        'Carrier': 'first',
        'Planned Ship Date': 'first'
    })
    
    # Reorder columns
    df = df[['Delivery Number', 'Company Name/Contact Name', 'Address 1', 'Address 2', 'Address 3',
             'City', 'State', 'Postal Code', 'Country', 'Product ID', 'Quantity', 'Sales Order',
             'Shipping Conditions', 'Delivery Instructions', 'Carrier', 'Planned Ship Date']]

    # Remove pandas index
    df = df.set_index('Delivery Number')
    
    return df

# Makes API calls to create orders in VeraCore
def create_orders(order_email : OrderEmail, orders: Orders, error_email: ErrorEmail):

    # Needs to be this to work
    headers = {
        "Content-Type" : "text/xml"
    }

    response = requests.post("https://rhu335.veracore.com/pmomsws/OMS.asmx", headers=headers, data=orders.generate_xml())

    if response.status_code > 299:
        # If error we want to add the offers to the error email
        error_email.add_offers(orders.offers)

        # Takes the error text and retrieves the relevant part adds that to body
        error_text = response.text
        split_string = error_text.split("System.Exception:")[-1]
        api_error = split_string.split("at")[0]
        error_email.add_to_body(orders.order_id, api_error)

        # Marks that there was an error and to send an email
        error_email.hasError = True
    else:
        # Or else add to the order email body to confirm what was uploaded
        order_email.add_to_body(orders.offers)


# Loads all environment variables
load_dotenv()

# All of this is found on Microsoft Entra under the application and the Overview tab

# Application Client ID
client_id = os.getenv("CLIENT_ID")
# Tenant ID for Microsoft Entra
tenant_id = os.getenv("TENANT_ID")
# Secret value generated in the application
client_secret = os.getenv("ENTRA_CLIENT_SECRET")
# Scope
scope = os.getenv("SCOPE")
# Authority
authority = f"https://login.microsoftonline.com/{tenant_id}"

# Resideo user id
resideo_id = os.getenv("USER")
# Resideo inbox id
inbox_id = os.getenv("INBOX_FOLDER")
# Resideo completed folder id
completed_id = os.getenv("COMPLETED_FOLDER")

# VeraCore Web User/Pass/System
veracore_id = os.getenv("VERACORE_USER")
veracore_pass = os.getenv("VERACORE_PASS")
veracore_system = os.getenv("VERACORE_SYSTEM")

# Generates an Outlook Draft email
def generate_outlook_email(user_id, email : Email, auth_header):
    generate_email_endpoint = f"https://graph.microsoft.com/v1.0/users/{user_id}/messages/"

    email_json = email.generate_email()

    response = requests.post(generate_email_endpoint, headers=auth_header, data=json.dumps(email_json))

    # If request is unsuccessful write to error log, otherwise return the draft id
    if not(response.status_code == 201):
        write_to_log(response.text)
        print(f"Draft wasn't created")
        return None
    else:
        print(f"Create draft : {response.status_code}")
        return response.json()["id"]

# Sends an Outlook draft
def send_outlook_email(user_id, draft_id, auth_header):
    send_draft_endpoint = f"https://graph.microsoft.com/v1.0/users/{user_id}/messages/{draft_id}/send"

    response = requests.post(send_draft_endpoint, headers=auth_header)

    # If the request isn't successful write to a log
    if not(response.status_code == 202):
        write_to_log(response.text)
        print(f"Email wasn't sent") 
    else:
        print(f"Send email : {response.status_code}")

# Moves an existing message in a mailbox
def move_outlook_email(user_id,email_id, endpoint_folder_id, auth_header):
    move_email_endpoint = f"https://graph.microsoft.com/v1.0/users/{user_id}/messages/{email_id}/move"

    move_body = {
        "destinationId" : endpoint_folder_id
    }

    response = requests.post(move_email_endpoint, headers=auth_header, data=json.dumps(move_body))

    # If request is not successful write to log
    if not(response.status_code == 201):
        write_to_log(response.text)
        print(f"Email wasn't moved")
    else:
        print(f"Move email : {response.status_code}")

# Creates a CSV attachment for a draft
def generate_attachment(user_id, email_id,csv_string, auth_header):
    attachment_endpoint = f"https://graph.microsoft.com/v1.0/users/{user_id}/messages/{email_id}/attachments"

    attachment_body = {
    "@odata.type": "#microsoft.graph.fileAttachment",
    "name": "ErrorOffer.csv",
    "contentType": "text/csv",
    "contentBytes": csv_string
    }

    response = requests.post(attachment_endpoint,headers=auth_header, data=json.dumps(attachment_body))

    # If request is not successful write to log
    if not(response.status_code == 201):
        write_to_log(response.text)
        print(f"Attachment wasn't created")
    else:
        print(f"Create attachment : {response.status_code}")


# Create a confidential application to verify with MSAL
app = ConfidentialClientApplication(
    client_id=client_id,
    client_credential = client_secret,
    authority=authority
)

# Get OAuth token for application
result = app.acquire_token_for_client(scopes=[scope])

# Create auth header
auth_header = {
    "Authorization" : f"Bearer {result["access_token"]}",
    "Content-Type" : "application/json"
}

# Get emails in the Resideo inbox
email_endpoint = f"https://graph.microsoft.com/v1.0/users/{resideo_id}/mailFolders/{inbox_id}/messages"

response = requests.get(email_endpoint, headers=auth_header)

# The value field holds the array of emails
emails = response.json()["value"]

for email in emails:
    # Only process emails with this in the subject
    if "REZISDC_OBD_" in email["subject"]:
        email_id = email["id"]
        # Get the attachment in the email
        attachment_endpoint = f"https://graph.microsoft.com/v1.0/users/{resideo_id}/messages/{email_id}/attachments"
        response = requests.get(attachment_endpoint, headers=auth_header)
        # Get the attachment ID to get the actual CSV
        attachment = response.json()["value"]
        attachment_id = attachment[0]["id"]

        # Get the actual bytes for the CSV
        value_endpoint = f"https://graph.microsoft.com/v1.0/users/{resideo_id}/messages/{email_id}/attachments/{attachment_id}/$value"

        response = requests.get(value_endpoint, headers=auth_header)

        # Read it into a pandas
        attachment_df = pd.read_csv(io.StringIO(response.text), sep="\t")

        # Process the df
        order_df = process_df(attachment_df)
        order_df = order_df.fillna("")
        order_df = order_df.sort_values(by="Delivery Number", ascending=True)

        # Create tuples from the dataframe
        order_tuples = order_df.itertuples()
        
        # Instantiate Email and Orders Objects
        email = OrderEmail()

        orders = Orders()
        
        error_email = ErrorEmail()

        # Loop through tuple to create order
        for order in order_tuples:

            # If the orders object is blank add order id
            if orders.order_id is None:
                orders.order_id = order[0]

            # If order IDs match add lines to the offers, otherwise send the API call and start on the next set of lines
            if orders.order_id == order[0]:
                orders.add_to_offers(order)
            else:
                
                create_orders(email,orders, error_email)

                orders = Orders(order[0])
                orders.add_to_offers(order)
        
        # One last call to get the last line
        create_orders(email,orders, error_email)
        
        # Generate the order drafts
        email_ids = []

        # If there is orders create an orders email
        if email.has_orders():
            email_ids.append(generate_outlook_email(resideo_id,email,auth_header))
        
        # If there is any errors create email and attach the missing lines
        if error_email.hasError:
            email_ids.append(generate_outlook_email(resideo_id, error_email, auth_header))
            generate_attachment(resideo_id,email_ids[-1],error_email.generate_error_bytes(),auth_header)

        # Send the draft
        for id in email_ids:
            
            if not(id == None):
                send_outlook_email(resideo_id,id,auth_header)

        # Move the email
        move_outlook_email(resideo_id,email_id,completed_id,auth_header)
    





        
        


    
    
    




            
    


