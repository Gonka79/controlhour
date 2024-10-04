# ControlHorario - Employee Time Tracking Bot

**ControlHorario** is a Python-based employee time tracking system built using a Telegram bot interface. This bot allows users to register their work hours by logging entry and exit times, making it ideal for managing and tracking employee attendance in small to medium-sized teams. All recorded data is stored in Excel files for easy access and analysis.

## Key Features:
- **User Registration**: New users can register their full name via the bot to identify their entries.
- **Entry and Exit Tracking**: Users can register their entry and exit times using simple commands (`/entrada` and `/salida`).
- **Data Storage**: All registered entries and exits are stored in an Excel file (`control_horarios.xlsx`), making it easy to maintain a history of attendance.
- **Hours Worked Calculation**: The bot automatically calculates the hours worked based on entry and exit times.
- **User Report**: Users can generate a summary report of their total hours worked with the `/reporte` command.

## How It Works:
1. **User Registration**: The user starts by registering their full name with the `/registrar` command.
2. **Register Entry**: To log their entry time, the user sends the `/entrada` command, and the bot records the current time as the entry.
3. **Register Exit**: The user can then send the `/salida` command to log their exit time. The bot calculates the total hours worked and updates the Excel file.
4. **Generate Report**: The `/reporte` command shows a summary of the user's total hours worked.

## Technologies Used:
- **Python**: Core programming language.
- **Telegram Bot API**: Used to interact with users via Telegram.
- **openpyxl**: Handles data storage and manipulation in Excel files.

## Commands:
- `/start`: Introduction and help command.
- `/registrar`: Register a new user with their full name.
- `/entrada`: Register entry time.
- `/salida`: Register exit time.
- `/reporte`: Generate a report of total hours worked.
- `/cancel`: Cancel any ongoing operation.

## Setup Instructions:
1. Clone the repository.
2. Install required dependencies:
   ```bash
   pip install -r requirements.txt

## RUN
```bash
python MerkaControl2.py(```)




## MIT License

Copyright (c) [Año] [Nombre del autor]

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.


