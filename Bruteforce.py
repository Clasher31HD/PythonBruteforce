import win32com.client as comclt
import time
wsh = comclt.Dispatch("WScript.Shell")
wsh.AppActivate("Trend Micro") # select another application

keys = ["a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z","A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","1","2","3","4","5","6","7","8","9","0"]
actual = 0

time.sleep(2)

current = keys[0]

for i in range(0, len(keys)):
    current = keys[actual]
    wsh.SendKeys(current)
    time.sleep(0.1)
    wsh.SendKeys("{ENTER}")
    time.sleep(0)
    wsh.SendKeys("{ENTER}")
    time.sleep(0.1)
    actual += 1
