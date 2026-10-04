# 🌍 Life Simulator

Life Simulator is a small interactive project where you can create your own virtual life and see how your daily choices affect different parts of it.

The idea is pretty simple — you start with a profile, some money, skills, energy and other stats, and then make decisions day by day.

It's basically a game + productivity simulator built using Python and Streamlit.

## 🚀 Live Demo

Try the app here:

https://lifesimulator.streamlit.app/

## 🎯 What is the idea?

In real life, small decisions can change things over time.

So I wanted to build something where those decisions could be simulated in a simple way.

You create your profile and then manage things like:

- Health
- Happiness
- Energy
- Skills
- Discipline
- Social life
- Money
- Goals
- Daily activities

As you progress, your stats change and your virtual life moves forward.

## ✨ Features

### 👤 Create Your Profile
Enter your basic information such as:

- Name
- Age
- Education
- Career
- Hobby
- Part-time job

### 📊 Life Stats

Keep track of your:

- ❤️ Health
- 😊 Happiness
- ⚡ Energy
- 🧠 Skills
- 🎯 Discipline
- 👥 Social
- 💰 Money

### 🎮 Daily Life Simulation

Each day gives you different choices and events.

Your decisions can increase or decrease your stats and affect your progress.

### 🎯 Goals

Create goals and track your progress instead of just playing randomly.

### 🏆 Achievements

The simulator has achievements that can be unlocked as you progress.

Some examples:

- One Week Survivor
- Level 5
- Skill Master
- Money Maker
- Discipline Master

### 💰 Money & Expenses

You can manage your virtual money and keep track of your transactions.

### 📅 Smart Timetable

The app can create a timetable based on your profile, career and current situation.

### 📜 Timeline

Your important actions and events are recorded so you can see how your virtual life has changed.

### 🤖 AI Feature

The project also includes a GenAI-based feature to make the experience more interactive and provide suggestions based on the user's situation.

## 🛠️ Tech Stack

- Python
- Streamlit
- OpenPyXL
- Google Gemini / GenAI
- Excel

## 💾 Data Storage

The project uses an Excel file to store the data.

Data such as profiles, goals, stats, achievements, timeline and money history can be saved so the information doesn't disappear every time the app is restarted.

## 📂 Project Structure

```text
Life-Simulator/
│
├── app.py
├── requirements.txt
├── Life_Simulator.xlsx
└── README.md
