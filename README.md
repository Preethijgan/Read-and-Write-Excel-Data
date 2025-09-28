# Aim:
To automate the process of reading data from an Excel file, processing it, and writing the output back into an Excel sheet using UiPath.

# Procedure:

1. Create Variables:
  - dtOrders (DataTable) rowscount (Int32) counter (Int32)
  <img width="1122" height="317" alt="Screenshot 2025-09-28 114134" src="https://github.com/user-attachments/assets/9b51631f-1a27-4206-849b-ec97752353f7" />

2. Add Excel Process Scope:
   - Drag Excel Process Scope to the Main Sequence.

3. Read Source File:
   - Inside the scope, add Use Excel File → select SourceData.xlsx.

4. Inside it, add Read Range:
   - Sheet Name = "Sheet1" (or actual sheet name) Output = dtOrders

5. Initialize Counters:
   - Add Assign: rowscount = dtOrders.Rows.Count Add Assign: counter = 0

6. Add Do While:
   - Condition: counter < rowscount
   - Inside:
        * Assign: dtOrders.Rows(counter)("Amount") = CInt(dtOrders.Rows(counter)("Amount")) + 10 Assign: counter = counter + 1

7. Write to Output File:
   - Add Use Excel File → select or type OutputData.xlsx and enable Create if not exists.

8. Inside it, add Write Range:
   - Input = dtOrders Sheet Name = "Orders" Start Cell = A1 Add Headers = Checked

9. Run the Workflow:
   - Execute the process. A new file OutputData.xlsx will be created with updated Amount values.

<img width="1919" height="963" alt="Screenshot 2025-09-28 114041" src="https://github.com/user-attachments/assets/b22f7aff-c0c5-4607-bb66-e55f58fb7e78" />
<img width="1919" height="1018" alt="Screenshot 2025-09-28 114059" src="https://github.com/user-attachments/assets/834f71b0-ed90-49e1-ae2d-a8cb7fc57551" />
<img width="1919" height="1019" alt="Screenshot 2025-09-28 114114" src="https://github.com/user-attachments/assets/acba9931-806d-49a6-9243-be13386bba9d" />



# Output:

## Sample data:
<img width="635" height="394" alt="Screenshot 2025-09-28 112911" src="https://github.com/user-attachments/assets/5191a247-fc7a-46c5-926c-a078476a9ff6" />


## After Read and Write:
<img width="621" height="335" alt="Screenshot 2025-09-28 113111" src="https://github.com/user-attachments/assets/6bee1018-390a-4a9c-86a9-d6438e826d7e" />





# Result

The workflow successfully reads the data from the given Excel sheet, processes it, and writes the updated/processed data into the desired Excel sheet.
