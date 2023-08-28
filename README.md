# Automatic YouTube Promotion Metrics
With just a simple copy and paste necessary, this Java program creates and formats an Excel document consisting of Cost per Sub & Impression to Sub Ratio columns, as well as automatically sorting based off performance. 

Additional Features:
- Promotion lists are separated by Status
- Total Active and Total counts are shown
- Rows are alternatively colored for easy readability
- Headings and List Titles are color coded for easy readability

![Sample Output](https://github.com/Ranchy101/YouTube-Promotion-Metrics/assets/42690717/e4e24120-4c2f-43ea-8a06-dedbae4f1b90)

Instructions:
1. Open your YouTube Promotions tab.
2. Scroll down to bottom and set Rows per Page to the maximum.
3. Copy and paste everything starting from "Promotion" all the way down to "Rows per page..."
4. Paste into promotions.txt file
5. Remove first 6 rows until you hit "Video thumbnail..."
6. Remove last 3 rows until you hit a number
7. Make sure previous output.xlsx are closed and that promotions.txt is saved in same directory
8. Compile and Run TestToExcel.java

Upcoming Changes:
- Automatic removal of "trash" rows in promotions.txt file, removing the need of steps 5-6.
  
