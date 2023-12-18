# Balbec_Case_Study
This repository encompasses my solution to the Balbec Capital Analyst case study, providing an analytical perspective on the data. To execute the code successfully, ensure that the Excel file is present in the same directory as the repository. You can achieve this by pulling the repository and manually placing the file into the shared folder.

While the Python implementation yielded satisfactory results compared to traditional Excel methods, certain observations indicate areas where Excel may be a more suitable choice for these specific requirements:

Formatting Challenges: After exporting data to the Excel spreadsheet, formatting alterations occurred, resulting in the loss of the original table outline. Though this doesn't impact the qualitative aspects, presenting this data to a client might necessitate additional effort to restore formatting.

Static Output: The code produces static output, requiring reruns whenever there is a modification to the original dataset. This contrasts with Excel, where dynamic updates are inherent.

Cell Traceability: Excel allows users to trace the source of data in a cell by clicking on the output of a formula. This feature is not available after Python output, potentially causing confusion if column headers lack clarity.

Percentage Calculation Accuracy: In Question 3, two-thirds of the percentages did not sum up to a precise 100, attributed to how the code accumulated sums. Given more time, implementing rounding techniques could enhance accuracy.

Note: For Question 2, two methods are presented in the Main.py file. While replicating the exact methodology of the original formula is challenging, the Python function demonstrates how a similar mistake could be made. Although this function is not invoked, its inclusion aims to illustrate my understanding of the second question's challenge.

Moreover, I opted to output the results to the same Excel file, a decision made for convenience, subject to modification.

Thank you for investing time in reviewing this case study. Should you have any inquiries, please feel free to contact me at thomasjditsworth@gmail.com.
