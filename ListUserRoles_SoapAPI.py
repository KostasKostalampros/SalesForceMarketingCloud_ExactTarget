'''
The following script connects to Exact Target (SalesForce Marketing Cloud) SOAP API
 and pulls all Account Users (AcountUser object) with the following properties:
 Name, AccountUserID, Email, ActiveFlag, LastSuccessfulLogin" , IsAPIUser, Roles, 
 DefaultBusinessUnit
'''

import requests
import xmltodict
import json
import csv

# Universal Variables
CLIENT_ID = "UPDATE_WITH_YOUR_CLIENT_ID"
CLIENT_SECRET = "UPDATE_WITH_YOUR_CLIENT_SECRET"

token_url = "https://auth.exacttargetapis.com/v1/requestToken"
data_url = "https://webservice.s4.exacttarget.com/Service.asmx"

output_csv_filename = "ExactTargetUserList.csv"


def make_request(request_url, request_payload, request_header):
    """Print input list of lists in a csv format in the same folder as this scripts.
    Args:
        request_url: string representing the Post request url.
        request_payload: string in xml format with the request parameters.
        request_header: dictionary including the Post request headers.
    Returns:
        token_response: dictionary including the Post request response 
    """
    token_response = requests.request("POST", request_url, data=request_payload, headers=request_header)

    return token_response


def print_table_to_csv(data_list, filename):
    """Print input list of lists in a csv format in the same folder as this scripts.
    Args:
        data_list: A list of lists populated with data.
        filename: A file name in string format used for the output file.
    Returns:
        None
    """
    with open(filename, "w", newline="") as file:
        writer = csv.writer(file, delimiter=",")
        for iter in range(len(data_list[0])):
            writer.writerow([x[iter] for x in data_list])


def main():
    token_payload = "{\n  \"clientId\": \"" + CLIENT_ID + "\",\n  \"clientSecret\": \"" + CLIENT_SECRET + "\"\n}"

    token_headers = {
        'Content-Type': "application/json",
        'Cache-Control': "no-cache",
    }

    data_headers = {
        'Content-Type': "text/xml",
        'SOAPAction': "Retrieve",
        'Cache-Control': "no-cache"
    }

    token_response = make_request(token_url, token_payload, token_headers)

    if token_response.status_code == 200:
        accessToken = token_response.json().get('accessToken')

        data_payload = "<soapenv:Envelope xmlns:soapenv=\"http://schemas.xmlsoap.org/soap/envelope/\" " \
                       "xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" " \
                       "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><soapenv:Header><fueloauth " \
                       "xmlns=\"http://exacttarget.com\">%s</fueloauth></soapenv:Header><soapenv:Body" \
                       "><RetrieveRequestMsg " \
                       "xmlns=\"http://exacttarget.com/wsdl/partnerAPI\"><RetrieveRequest><ObjectType>AccountUser" \
                       "</ObjectType><Properties>Name</Properties><Properties>AccountUserID</Properties><Properties" \
                       ">Email</Properties><Properties>ActiveFlag</Properties><Properties>LastSuccessfulLogin" \
                       "</Properties><Properties>IsAPIUser</Properties><Properties>Roles</Properties><Properties" \
                       ">DefaultBusinessUnit</Properties></RetrieveRequest></RetrieveRequestMsg></soapenv:Body" \
                       "></soapenv:Envelope>" % accessToken

        data_response = make_request(data_url, data_payload, data_headers)
        response_text = data_response.text

        # Prints Exact Target response data into a xml file for data exploration
        # with open('response_output.txt', 'w', newline="") as opened_file:
        # opened_file.write(response_text)

        xmlUserData = xmltodict.parse(response_text)
        jsonUserData = json.dumps(xmlUserData)
        dictUserData = json.loads(jsonUserData)

        # Prints interim data into a json file with pretty format for data exploration
        # with open('json_output.json', 'w', newline="") as opened_file:
        # opened_file.write(json.dumps(dictUserData, indent=4))

        results = dictUserData.get('soap:Envelope').get('soap:Body').get('RetrieveResponseMsg').get('Results')

        list_userId = ['userId']
        list_userName = ['userName']
        list_userEmail = ['userEmail']
        list_status = ['status']
        list_lastSuccessfulLogin = ['lastSuccessfulLogin']
        list_isAPIUser = ['isAPIUser']
        list_roles = ['roles']
        list_defaultBusinessUnit = ['defaultBusinessUnit']

        # Parse all AccountUser objects and collect information about the user
        for result in results:

            userId = result.get('AccountUserID')
            userName = result.get('Name')
            userEmail = result.get('Email')
            status = result.get('ActiveFlag')
            lastSuccessfulLogin = result.get('LastSuccessfulLogin')
            isAPIUser = result.get('IsAPIUser')
            defaultBusinessUnit = result.get('DefaultBusinessUnit')

            json_roles = result.get('Roles').get('Role')
            userRoles = ""
            for json_role in json_roles:
                if type(json_role) is dict:
                    userRoles = userRoles + json_role.get('Name') + ';'

            list_userId.append(userId)
            list_userName.append(userName)
            list_userEmail.append(userEmail)
            list_status.append(status)
            list_lastSuccessfulLogin.append(lastSuccessfulLogin)
            list_isAPIUser.append(isAPIUser)
            list_defaultBusinessUnit.append(defaultBusinessUnit)
            list_roles.append(userRoles)

        final_user_data_list = [list_userId, list_userName, list_userEmail, list_status, list_lastSuccessfulLogin,
                                list_isAPIUser, list_defaultBusinessUnit, list_roles]

        print_table_to_csv(final_user_data_list, output_csv_filename)


if __name__ == '__main__':
    main()
