from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.textinput import TextInput
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.anchorlayout import AnchorLayout
from kivy.uix.gridlayout import GridLayout
from kivy.graphics import Color
import os

import pandas as pd
import tkinter as tk
from tkinter import filedialog

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


class LoginScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        anchor_layout = AnchorLayout(anchor_x='center', anchor_y='center')  # Center-align the GridLayout
        layout = GridLayout(cols=2, spacing=10, size_hint=(0.6, 0.3))
        self.email_label = Label(text="Email:")
        self.email_input = TextInput()
        self.password_label = Label(text="Password:")
        self.password_input = TextInput(password=True)
        self.login_button = Button(text="Login", on_press=self.on_login)

        #layout = BoxLayout(orientation='vertical', spacing=10)
        layout.add_widget(self.email_label)
        layout.add_widget(self.email_input)
        layout.add_widget(self.password_label)
        layout.add_widget(self.password_input)
        layout.add_widget(self.login_button)
        anchor_layout.add_widget(layout)
        self.add_widget(anchor_layout)

    def on_login(self, instance):
        # Implement login functionality here
        #self.manager.current = 'compose'
        email = self.email_input.text
        password =self.password_input.text
        print(email,password)
        app.email = email
        app.password = password
         
        app.root.current = 'compose'
        app.root.get_screen('compose')
        
        
        
        
        #print(email,password)

class ComposeScreen(Screen):
    def __init__(self,**kwargs):
        super().__init__(**kwargs)
        
        self.file_label = Label(text="choose File")
        self.file_button = Button(text="Browser",on_press=self.file)

        
        self.receiver_label = Label(text="Receiver Email:")
        self.receiver_input = TextInput()
        self.total_label = Label(text="Total  Email Found:")
        self.total_input = TextInput()
        
        self.subject_label = Label(text="Subject:")
        self.subject_input = TextInput(hint_text='Subject')
        self.message_label = Label(text="Message:",width=150)
        self.message_input = TextInput(hint_text='Message')
        self.send_button = Button(text="Send",background_color=(0, 1, 0, 1), on_press=self.on_send)
        self.Reset_button = Button(text="Reset", on_press=self.reset)
        self.label = Label(text='', size_hint_y=None)
        self.Back_button = Button(text='Back', on_press=self.back)
        anchor_layout = AnchorLayout(anchor_x='center', anchor_y='center')  # Center-align the GridLayout
        layout = GridLayout(cols=2, spacing=10, size_hint=(0.8, 0.6),width=200)

        #layout = BoxLayout(orientation='vertical', spacing=10)
        layout.add_widget(self.file_label)
        layout.add_widget(self.file_button)
        
        layout.add_widget(self.receiver_label)
        layout.add_widget(self.receiver_input)
        layout.add_widget(self.total_label)
        layout.add_widget(self.total_input)
        #self.add_widget(self.file_chooser)
        
        layout.add_widget(self.subject_label)
        layout.add_widget(self.subject_input)
        layout.add_widget(self.message_label)
        layout.add_widget(self.message_input)
        layout.add_widget(self.send_button)
        layout.add_widget(self.Reset_button)
        
        layout.add_widget(self.label)
        layout.add_widget(self.Back_button)
        #self.add_widget(layout)
        anchor_layout.add_widget(layout)
        self.add_widget(anchor_layout)
        
    def back(self,instance):
        app.root.current = 'login'
        app.root.get_screen('login')
        
        
 
        
        pass
        
    def file(self,instance):
        file_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", "*.xlsx")],
        title="Select an Excel file"
    )    
        df = pd.read_excel(file_path, sheet_name='Sheet1')
        
        # Extract the email addresses from the 'Emails' column and join them with commas
        email_ids = ', '.join(df['email'].astype(str))
        
        print(email_ids)
        
        
        
        self.receiver_input.text = f'{email_ids}'
        email_id = self.receiver_input.text.split(',')

        # Count the number of email addresses in the list
        count = len(email_id)
        self.total_input.text = f'{count}'
        
        print(count)
        

    
        


    def on_send(self, instance):
        # Implement email sending functionality here
        email = app.email
        password=app.password
        
        print("Email id : ",f'{email}')
        print("password: ",f'{password}')
        
        #mail send process
        
        recipients = self.receiver_input.text.split(',')
        subject = self.subject_input.text
        message_text = self.message_input.text
        
        # Connect to the SMTP server
        smtp_server = 'smtp.gmail.com'  # Use Gmail SMTP server
        smtp_port = 587
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # Enable TLS

        # Login to your email account
        server.login(email, password)
        self.label.text = "Please wait few second"
        # Send the email to each recipient
        for recipient in recipients:
           
            msg = MIMEMultipart()
            msg['From'] = email
            msg['To'] = recipient.strip()  # Set the "To" header for the current recipient
            msg['Subject'] = subject

            # Add the message text
            msg.attach(MIMEText(message_text, 'plain'))

            server.sendmail(email, recipient, msg.as_string())
            self.label.text = "mail sent"

        # Quit the SMTP server
        server.quit()

        pass
    def reset(self, instance):
        
        self.receiver_input.text=""
        self.message_input.text= ""
        self.subject_input.text= " "
        self.total_input.text= " "
        self.label.text = " "
        pass
class EmailApp(App):
    email =''
    password =''
    def build(self):
        self.title="TakeOff"
        self.screen_manager = ScreenManager()

        login_screen = LoginScreen(name='login')
        compose_screen = ComposeScreen(name='compose')

        self.screen_manager.add_widget(login_screen)
        self.screen_manager.add_widget(compose_screen)

        return self.screen_manager

if __name__ == '__main__':
    app = EmailApp()
    app.run()