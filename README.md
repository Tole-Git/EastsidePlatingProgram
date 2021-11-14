# EastsidePlatingProgram
Est. June 2020
Author: Tony Le

This project is my first project I've ever made for a company, and also my first ever project practicing skills learned
from school and applying them to a real-world setting. 

Apache POI must be downloaded and imported in order for this program to be used. 

This project consists of a program that takes in inventory data in order to send to plating companies to process.
One important information that must be utilized are weight per part. Since boxes being shipped to plating companies come
in all sorts of different sizes and numbers, a lot of time is wasted by weighing each box. In order to combat this and 
save more than 5+ working hours per week, the solution is below. 

The feature that defines this program is the varifyExcel() method. This method tells the user if the entered part-number has been
entered before. This is specifically useful because instead of normally weighing the box out again, we can use previous 
"weight-per-part" data to construct the final weight of that box based on the amount of parts given. This not only saves time and
money, box also the use of excess force since boxes can weigh upwards to 50+ pounds.
