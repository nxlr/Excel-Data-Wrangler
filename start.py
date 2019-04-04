import os
import time

#start = time.time()

os.system("python gva_current_prices.py")
os.system("python gva_constant_prices.py")
os.system("python nva_current_prices.py")
os.system("python nva_constant_prices.py")

#end = time.time()

#exec_time = end - start
#print("\n Execution Time : ", exec_time, "\n")