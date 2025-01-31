# Timetable Bot 🗓️🤖

This is a Telegram bot that helps users generate personalized timetables based on their course selections. The bot interacts with users to collect course information, processes timetable data from CSV files, and generates an Excel file containing the user's timetable.

---

## Features

- **Telegram Bot Interaction**: Users can interact with the bot to input their courses and receive a timetable.
- **Course Search**: Searches for user-provided courses in predefined timetable CSV files.
- **Timetable Generation**: Generates a pivot table for the timetable and exports it to an Excel file.
- **Excel Formatting**: Formats the Excel file with bold headers, wrapped text, and centered alignment.
- **CSV Integration**: Reads timetable data from CSV files for each day of the week.

---

## Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/your-username/timetable-bot.git
   cd timetable-bot
   ```

2. **Set up a virtual environment** (optional but recommended):
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Set up Telegram Bot Token**:
   Replace the placeholder token in the code with your actual Telegram bot token:
   ```python
   bot = telebot.TeleBot("YOUR_TELEGRAM_BOT_TOKEN")
   ```

5. **Run the bot**:
   ```bash
   python app.py
   ```

---

## How It Works

1. **Start the Bot**:
   - Use the `/start` command to begin interacting with the bot.

2. **Input Courses**:
   - Use the `/courses` command to input your courses.
   - Enter the number of courses you want to add.
   - Provide the course codes one by one (e.g., `ECN311`).

3. **Generate Timetable**:
   - The bot processes the input courses and searches for them in the timetable CSV files.
   - It generates a timetable in Excel format and sends it to the user.

---

## File Structure

```
timetable-bot/
├── app.py                # Main bot application
├── timetable/            # Directory containing timetable CSV files
│   ├── monday.csv
│   ├── tuesday.csv
│   ├── wednesday.csv
│   ├── thursday.csv
│   └── friday.csv
├── requirements.txt      # List of dependencies
└── README.md             # Project documentation
```

---

## Dependencies

- Python 3.8+
- `pandas` for data manipulation
- `openpyxl` for Excel file generation and formatting
- `matplotlib` for visualization (not actively used in the current implementation)
- `telebot` for Telegram bot interaction

---

## Contributing

Contributions are welcome! If you'd like to contribute to the Timetable Bot, please follow these steps:

1. Fork the repository.
2. Create a new branch for your feature or bugfix.
3. Commit your changes.
4. Submit a pull request.

---

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

---

Generate your timetable with ease using the Timetable Bot! 🚀📅
