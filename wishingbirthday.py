import pandas as pd
import datetime
import smtplib


GMAIL_ID = 'jaiarpit123@gmail.com'
GMAIL_PASSWORD = 'ddad dmal lrkp hbua'

def sendEmail(to, sub, msg):
    # print(f"Email to {to} send with subject: {sub} and message {msg}")

    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PASSWORD)
    s.sendmail(GMAIL_ID, to, f"Subject: {sub}\n\n{msg}")
    s.quit()

# sendEmail("jaiarpit123@gmail.com", "Hey", "Hi")
# exit()

if __name__=='__main__':
    df = pd.read_excel("databirth.xlsx")
    # print(df)

    today = datetime.datetime.now().strftime("%d-%m")
    yearnow = datetime.datetime.now().strftime("%Y")
    # print(today, yearnow)

    writeInd = []
    for index, item in df.iterrows():
        # print(index, item["Birthday"])
        bday = item["Birthday"].strftime("%d-%m")
        # print(bday)

        if today == bday and yearnow not in str(item["Year"]):
            sendEmail(item["Email"], "Happy Birthday", item["Dialogue"])
            writeInd.append(index)

    # print(writeInd)
    if writeInd:
        for i in writeInd:
            yr = df.loc[i, "Year"]
            # print(yr)
            df.loc[i, "Year"] = f"{str(yr)}, {str(yearnow)}"
            # print(df.loc[i, "Year"])

        print(df)
        df.to_excel("databirth.xlsx", index=False)