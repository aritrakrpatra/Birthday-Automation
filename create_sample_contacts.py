"""One-off helper to create a sample contacts.xlsx file to get started."""

import pandas as pd

sample_data = [
    {
        "Name": "Alice Johnson",
        "Email": "alice@example.com",
        "Birthday": "04-10-1995",
        "Subject": "",
        "Personalised Message": "",
    },
    {
        "Name": "Bob Smith",
        "Email": "bob@example.com",
        "Birthday": "15-08-1990",
        "Subject": "Happy Birthday, {name}! 🎉 - From Cholo Jai Dure Tour & Travels",
        "Personalised Message": (
            "Hey {name},\n\n"
            "Happy Birthday from Cholo Jai Dure Tour & Travels! Hope your special day "
            "is as wonderful as your next trip with us.\n\n"
            "Cheers,\nTeam Cholo Jai Dure Tour & Travels"
        ),
    },
]

if __name__ == "__main__":
    df = pd.DataFrame(sample_data)
    df.to_excel("contacts.xlsx", index=False)
    print("Created contacts.xlsx with sample data.")
