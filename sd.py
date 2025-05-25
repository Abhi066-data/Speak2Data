import pandas as pd
import matplotlib.pyplot as plt
import speech_recognition as sr
import pyttsx3
import re
import tkinter as tk
from tkinter import messagebox
import os

# Load and preprocess data
df = pd.read_csv(r"D:\.vscode\Folder\samp.csv")

# Normalize string columns globally
for col in df.select_dtypes(include='object').columns:
    df[col] = df[col].astype(str).str.strip().str.lower()

# Path to user's documents folder (for saving files)
documents_path = os.path.expanduser("~/Documents")

# Text-to-speech engine
def speak(text):
    engine = pyttsx3.init()
    engine.say(text)
    engine.runAndWait()

# Voice input
def get_voice_command():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        status_label.config(text="🎙️ Listening...")
        listen_button.config(state=tk.DISABLED, bg="#FF6347")
        audio = r.listen(source)
    try:
        command = r.recognize_google(audio)
        print(f"🗣️ You said: {command}")
        status_label.config(text="✅ Command Received!")
        speak(command)
        process_command(command.lower())
    except Exception as e:
        print(e)
        speak("Sorry, I didn't get that.")
        status_label.config(text="❌ Could not understand.")
    finally:
        listen_button.config(state=tk.NORMAL, bg="#4CAF50")

# Data cleaning logic
def clean_data(command):
    global df
    performed = False

    if ("fill" in command and ("missing" in command or "null" in command) and "mean" in command) \
       or "fill missing" in command or "fill null" in command:
        speak("Filled missing values with the mean.")
        performed = True

    if "remove duplicates" in command:
        df = df.drop_duplicates()
        speak("Removed duplicate rows.")
        performed = True

    if "trim whitespace" in command:
        for col in df.select_dtypes(include='object').columns:
            df[col] = df[col].str.strip()
        speak("Trimmed whitespace in text fields.")
        performed = True

    if "remove sales null" in command:
        if 'Sales' in df.columns:
            df = df[df['Sales'].notnull()]
            speak("Removed rows with NaN sales values.")
            performed = True
        else:
            speak("Sales column not found.")

    if "remove invalid" in command:
        if 'Price' in df.columns:
            df = df[df['Price'] >= 0]
            speak("Removed rows with invalid price values.")
            performed = True
        else:
            speak("Price column not found.")

    if performed:
        status_label.config(text="✅ Data cleaned.")
    else:
        speak("I didn't understand the data cleaning command.")
        status_label.config(text="❌ Invalid cleaning command.")

# Process command function
def process_command(command):
    global df

    if "open excel" in command:
        try:
            status_label.config(text="✅ Excel opened.")
        except Exception as e:
            speak("Failed to open Excel.")
            print(e)
        return

    if any(keyword in command for keyword in [
        "clean data", "remove missing", "remove null", "fill missing", "fill null",
        "remove duplicates", "trim whitespace", "remove invalid", "remove sales null"
    ]):
        clean_data(command)
        return

    # Filtering logic
    if "show" in command or "filter" in command:
        year = None
    
        if year_match:
            year = int(year_match.group(1))

        if "Name" in df.columns:
            possible_names = df["Name"].dropna().unique()
            if matched_names:
                name = matched_names[0]

        numbers = re.findall(r"\b(\d{4,6})\b", command)
        if numbers:
            for num in numbers:
                n = int(num)
                if not year or n != year:
                    price = n
                    break

        temp_df = df.copy()
        applied_filters = []

        if year:
            if "Year" in temp_df.columns:
                temp_df = temp_df[temp_df["Year"] == year]
                applied_filters.append(f"Year: {year}")
            else:
                speak("Year column not found.")

        if name:
            if "Name" in temp_df.columns:
                applied_filters.append(f"Name: {name}")
            else:
                speak("Name column not found.")

        if price:
            if "Price" in temp_df.columns:
                temp_df = temp_df[temp_df["Price"] >= price]
                applied_filters.append(f"Price >= {price}")
            else:
                speak("Price column not found.")

        if temp_df.empty:
            speak("No matching data found.")
            status_label.config(text="⚠️ No matching data found.")
        else:
            print(temp_df)
            temp_df.to_excel("filtered_output.xlsx", index=False)
            speak(f"Data filtered by {', '.join(applied_filters)}. Exported to Excel.")
            status_label.config(text="✅ Filtered & Exported!")
        return

    if any(keyword in command for keyword in ["show null", "check null", "null rows", "null columns"]):
        null_rows = df[df.isnull().any(axis=1)]
        null_columns = df.loc[:, df.isnull().any()]

        if not null_rows.empty:
            null_rows.to_excel("null_rows.xlsx", index=False)
            speak("Exported rows with null values to Excel.")
            status_label.config(text="✅ Null rows exported.")
        else:
            speak("No rows with null values found.")
            status_label.config(text="✅ No null rows.")

        if not null_columns.empty:
            null_columns.to_excel("null_columns.xlsx", index=False)
            speak("Exported columns with null values to Excel.")
        else:
            speak("No columns with null values found.")
        return

    if "count null" in command or "count nan" in command:
        null_summary = df.isnull 
        null_summary.to_excel("null_summary.xlsx")
        speak(f"There are {total_nulls} missing values. Exported summary to Excel.")
        return

    if "compare" in command:
        if "Name" not in df.columns or "Revenue" not in df.columns:
            speak("Name or Revenue column not found.")
            return

        possible_names = df["Name"].unique()
        words = command.lower().split()
        names = [word for word in words if word in possible_names]

        if len(names) >= 2:
            comp_df = df[df["Name"].isin(names)]
            comp_df.groupby("Name")["Revenue"].sum().plot(kind="bar")
            plt.title("Revenue Comparison")
            plt.xlabel("Name")
            plt.ylabel("Revenue")
            plt.tight_layout()
            plt.show()
            speak("Here is the comparison chart.")
        else:
            speak("Please mention at least two valid names to compare.")
        return

    if "plot" in command or "chart" in command:
        if "category" in command and "sales" in command:
            if "Category" in df.columns and "Sales" in df.columns:
                df.groupby("Category")["Sales"].sum().plot(kind="bar", color="skyblue")
                plt.title("Category-wise Sales")
                plt.xlabel("Category")
                plt.ylabel("Sales")
                plt.tight_layout()
                plt.show()
                speak("Here is the sales chart.")
            else:
                speak("Category or Sales column not found.")
        elif "year" in command and "revenue" in command:
            if "Year" in df.columns and "Revenue" in df.columns:
                df.groupby("Year")["Revenue"].sum().plot(kind="line", marker="o")
                plt.title("Yearly Revenue Trend")
                plt.xlabel("Year")
                plt.ylabel("Revenue")
                plt.grid(True)
                plt.tight_layout()
                plt.show()
                speak("Here is the revenue trend.")
            else:
                speak("Year or Revenue column not found.")
        else:
            speak("Sorry, I couldn't understand the chart request.")
            status_label.config(text="❌ Invalid chart request.")
        return

    speak("I didn't understand the request.")
    status_label.config(text="❌ Could not understand.")

# Button click logic
def on_button_click():
    speak("Listening for your command...")

# GUI layout
root = tk.Tk()
root.title("Speak2Data")
root.geometry("500x300")
root.configure(bg="#f4f4f4")

title = tk.Label(root, text="🎙️ Voice Analyst", font=("Arial", 20, "bold"), bg="#f4f4f4", fg="#333")
title.pack(pady=20)

listen_button = tk.Button(
    root, text="Tap to Speak", command=on_button_click, font=("Arial", 14),
    bg="#4CAF50", fg="white", relief="solid", bd=2, width=20, height=2, activebackground="#45a049"
)
listen_button.pack(pady=20)
status_label = tk.Label(root, text="Click to start listening.", font=("Arial", 14), bg="#f4f4f4", fg="#333")
status_label.pack(pady=10)

root.mainloop()
