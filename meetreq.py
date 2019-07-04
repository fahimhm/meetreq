import time
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(1)
mail.Start = '2019-08-28 17:00'
mail.Subject = 'Test subject'
mail.Duration = 15
mail.Location = 'Meeting Location'
mail.MeetingStatus = 1
mail.Recipients.Add("fahimhadimaula@gmail.com")
mail.Save()
mail.Send()