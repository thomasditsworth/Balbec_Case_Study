# Balbec_Case_Study
This repository encompasses my solution to the Balbec Capital Analyst case study, providing an analytical perspective on the data. To execute the code successfully, ensure that the Excel file is present in the same directory as the repository. You can achieve this by pulling the repository and manually placing the file into the shared folder.

While the Python implementation gave satisfactory results compared to traditional Excel methods, I noticed a couple of areas where Excel may be a more suitable choice for these specific requirements. Given more advanced requirements, python could be incredilby useful, and I really enjoyed gaining more expereience in this particular line of coding. 

Formatting Challenges: After exporting data to the Excel spreadsheet, formatting alterations occurred, resulting in the loss of the original table outline. Though this doesn't impact the qualitative aspects, presenting this data to a client might necessitate additional effort to restore formatting. As with the percentage calculations, there are several more advanced ways to potentially maintain some of the table strucutre, that I would definetly be keen on diving deeper into. 

Static Output: The code produces static output, requiring reruns whenever there is a modification to the original dataset. This contrasts with Excel, where dynamic updates make it a lot more intutive. If there are no headers or the data is unclear, this could cause some issues. 

Cell Traceability: Excel allows users to trace the source of data in a cell by clicking on the output of a formula. This feature is not available after Python output, potentially causing confusion if column headers lack clarity.

Percentage Calculation Accuracy: In Question 3, two-thirds of the percentages did not sum up to a precise 100, attributed to how the code accumulated sums. Given more time, implementing rounding techniques could enhance accuracy.

Note: For Question 2, two methods are presented in the Main.py file. While replicating the exact methodology of the original formula is challenging, the Python function demonstrates how a similar mistake could be made. Although this function is not invoked, its inclusion aims to illustrate my understanding of the second question's challenge.

Moreover, I opted to output the results to the same Excel file, a decision made for convenience, and is easily subject to modification.

Thank you so much for taking the time to review this case study. If you have any questions please feel free to contact me at thomasjditsworth@gmail.com.
