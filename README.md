# Loan Amortization Schedule VBA Add-in

## Overview
The **Loan Amortization Schedule VBA Add-in** is an Excel add-in designed to simplify the creation of detailed loan amortization schedules. It allows users to input loan parameters such as principal, interest rate, loan term, and tenure to generate a clear and structured repayment schedule.

## Features
- Easy-to-use interface for entering loan details.
- Generates a comprehensive loan amortization schedule.
- Supports fixed interest rates.
- Options for only monthly payment schedules.
- Displays key metrics such as total interest paid and remaining balance.
- Supports early repayment and balloon payments.

## Installation
1. Download the `.xlam` file for the add-in from ([here](https://github.com/DanishDanial-Vba/Excel-Amortization-Schedule/raw/main/Finance%20Lease.xlam)
).
2. Open Excel and go to `File > Options > Add-ins`.
3. At the bottom of the window, select `Excel Add-ins` in the "Manage" dropdown and click **Go**.
4. Click **Browse** and select the downloaded `.xlam` file.
5. Check the box next to the add-in name to enable it.
6. Click **OK** to load the add-in.

## Usage
1. After enabling the add-in, navigate to the **Finance** tab that appears on the Excel ribbon.
2. Click **Finance Lease** to open the input form.
3. Enter the following loan details:
   - **Principal Amount**: Total loan amount.
   - **Loan Term**: Duration of the loan in months.   
   - **Optional Inputs**: Early payments or balloon payments.
   - **Option Choice**: Choose an option based on what information you have.   
   - **Interest Rate/Installment amount**: Annual percentage rate (APR) or the Installment amount based on the option selected.
4. Click **Compute**: Compute the Interest or installment based on the option selected.
5. Choose the **Loan start date**:
6. **Installment Start Date**: is frozen and is set to one month after the loan start date
7. Click **Generate Schedule** to produce the amortization table in a new worksheet.

## Example Output
The amortization schedule includes:
- Date
- Opening Balance
- Days
- Interest
- Service Fee   
- Installment Amount
- Principal Paid   
- Interest Paid
- Service Fee Paid
- Closing Balance

## Requirements
- Microsoft Excel 2013 or later (Windows or Mac).
- VBA macros must be enabled.

## Licensing
This add-in is licensed under the **[GPL-3.0 License](https://www.gnu.org/licenses/gpl-3.0.html)**. You are free to use, modify, and distribute this add-in under the terms of this license.

## Contributing
We welcome contributions from the community! If you'd like to report a bug, request a feature, or contribute to the project:
1. Fork the repository.
2. Create a feature branch (`git checkout -b feature-name`).
3. Commit your changes (`git commit -m "Added feature"`).
4. Push to the branch (`git push origin feature-name`).
5. Open a pull request.

## Support
If you encounter any issues, please open an [issue](#) or contact us at `danish.ukn@gmail.com`.

## Credits
Developed by Danish Danial.
