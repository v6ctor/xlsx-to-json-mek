# AYCD Shopify profile converter alternative for MEK AIO

from python_calamine import CalamineWorkbook

import sys
import json
import uuid
import os

argv_length = len(sys.argv)

if (argv_length != 3 or ".xlsx" not in sys.argv[1]):
    print("usage: python3 profile-convert.py <import_file_name>.xlsx <export_file_name>")
    quit()

us_state_to_abbrev = {
    "alabama": "AL",
    "alaska": "AK",
    "arizona": "AZ",
    "arkansas": "AR",
    "california": "CA",
    "colorado": "CO",
    "connecticut": "CT",
    "delaware": "DE",
    "florida": "FL",
    "georgia": "GA",
    "hawaii": "HI",
    "idaho": "ID",
    "illinois": "IL",
    "indiana": "IN",
    "iowa": "IA",
    "kansas": "KS",
    "kentucky": "KY",
    "louisiana": "LA",
    "maine": "ME",
    "maryland": "MD",
    "massachusetts": "MA",
    "michigan": "MI",
    "minnesota": "MN",
    "mississippi": "MS",
    "missouri": "MO",
    "montana": "MT",
    "nebraska": "NE",
    "nevada": "NV",
    "new hampshire": "NH",
    "new jersey": "NJ",
    "new mexico": "NM",
    "new york": "NY",
    "north carolina": "NC",
    "north dakota": "ND",
    "ohio": "OH",
    "oklahoma": "OK",
    "oregon": "OR",
    "pennsylvania": "PA",
    "rhode island": "RI",
    "south carolina": "SC",
    "south dakota": "SD",
    "tennessee": "TN",
    "texas": "TX",
    "utah": "UT",
    "vermont": "VT",
    "virginia": "VA",
    "washington": "WA",
    "west virginia": "WV",
    "wisconsin": "WI",
    "wyoming": "WY",
    "district of columbia": "DC",
    "american samoa": "AS",
    "guam": "GU",
    "northern mariana islands": "MP",
    "puerto rico": "PR",
    "united states minor outlying islands": "UM",
    "u.s. virgin islands": "VI",
}

abbrev_states_keys = list(us_state_to_abbrev.keys())
abbrev_states_values = list(us_state_to_abbrev.values())

workbook = CalamineWorkbook.from_path(sys.argv[1]).get_sheet_by_index(0).to_python(skip_empty_area=False)

dictionary = {
    "aioExportProfiles": []
}

count = 1
for user in workbook:
    if (count == 1):
        count = 0
        continue
    
    shipping_name = user[24].split()
    shipping_first_name = shipping_name[0]
    shipping_last_name = ' '.join(shipping_name[1:])

    if (user[31] == "California "):
        user[31] = "california"

    if (len(user[31].strip()) == 2):
        shipping_state_long = abbrev_states_keys[abbrev_states_values.index(user[31].upper())].title()
        shipping_state_short = abbrev_states_values[abbrev_states_values.index(user[31].upper())]

    else:

        shipping_state_long = user[31].lower().title()
        shipping_state_short = us_state_to_abbrev.get(abbrev_states_keys[abbrev_states_keys.index(user[31].lower())]).upper()

    if (len(user[22].strip()) == 2):
        billing_state_long = abbrev_states_keys[abbrev_states_values.index(user[22].upper())].title()
        billing_state_short = abbrev_states_values[abbrev_states_values.index(user[22].upper())]

    else:
        billing_state_long = user[22].lower().title()
        billing_state_short = us_state_to_abbrev.get(abbrev_states_keys[abbrev_states_keys.index(user[31].lower())]).upper()

    billing_name = user[15].split()
    billing_first_name = billing_name[0]
    billing_last_name = ' '.join(billing_name[1:])

    shipping_information = {
        "firstName": shipping_first_name,
            "lastName": shipping_last_name,
            "address1": user[26],
            "address2": user[27],
            "state": {
            "short": shipping_state_short,
            "long": shipping_state_long
            },
            "city": user[30].lower().title(),
            "country": {
            "short": "US",
            "long": "United States"
            },
            "zip": user[29]
    }

    billing_information = {
        "firstName": billing_first_name,
            "lastName": billing_last_name,
            "address1": user[17],
            "address2": user[18],
            "state": {
            "short": billing_state_short,
            "long": billing_state_long
            },
            "city": user[21].lower().title(),
            "country": {
            "short": "US",
            "long": "United States"
            },
            "zip": user[20]
    }

    if (user[14] == "Yes"):
        billing_information = shipping_information

    dictionary.get("aioExportProfiles").append(
        {
            "name": user[6],
            "email": user[5],
            "phone": user[25],
            "cardHolder": user[8],
            "cardNum": user[10],
            "cvv": user[13],
            "expmonth": user[11],
            "expyear": user[12],
            "isQuickProfile": False,
            "isOneCheckout": user[7] == "True",
            "isBillingSameAsShipping": user[14] == "Yes",
            "shipping": shipping_information,
            "billing": billing_information,
            "key": str(uuid.uuid4())
        }
    )

with open(sys.argv[2] + ".json", 'w') as export:
    export.write(json.dumps(dictionary, indent=4))

print("Successfully exported to file: %s" % os.path.abspath(sys.argv[2]))



