import requests
import wikipedia
import pywhatkit as kit
from email.message import EmailMessage
import smtplib
from decouple import config

EMAIL = "shamukh.sriharsha@gmail.com"
PASSWORD = "zbzd tgph doxn ezlf"


def find_my_ip():
    ip_address = requests.get('https://api.ipify.org?format=json').json()
    return ip_address["ip"]


def search_on_wikipedia(query):
    results = wikipedia.summary(query, sentences=2)
    return results


def search_on_google(query):
    kit.search(query)


def youtube(video):
    kit.playonyt(video)

def send_email(receiver_add, subject, message):
    try:
        email = EmailMessage()
        email['To'] = receiver_add
        email['Subject'] = subject
        email['From'] = EMAIL

        email.set_content(message)
        s = smtplib.SMTP("smtp.gmail.com", 587)
        s.starttls()
        s.login(EMAIL, PASSWORD)
        s.send_message(email)
        s.close()
        return True


    except Exception as e:
        print(e)
        return False


def get_news():
    news_headline = []
    result = requests.get(f"https://newsapi.org/v2/top-headlines?country=in&category=general&apiKey"
                          f"=4d67d71efa2d4fbba24495c04de5f61d").json()
    articles = result["articles"]
    for article in articles:
        news_headline.append(article['title'])
    return news_headline[:6]


import requests
def weather_forecast(city):
    api_key = "2aa57600c894dd2fbf3227614007877d"
    try:
        res = requests.get(f"https://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}&units=metric")
        res.raise_for_status()
        data = res.json()
        weather = data["weather"][0]["main"]
        temp = data["main"]["temp"]
        feels_like = data["main"]["feels_like"]

        return weather, f"{temp}°C", f"{feels_like}°C"
    except requests.exceptions.RequestException as e:
        print(f"Error fetching weather data: {e}")
        return None, None, None

