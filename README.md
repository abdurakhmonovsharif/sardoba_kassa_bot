# Sardoba Restaurant Telegram Bot

## Docker

### Sardoba server uchun
`.env` ichida kamida shu qiymatlar bo'lsin:
- `TELEGRAM_BOT_TOKEN=...`
- `DATABASE_URL=postgresql://user:password@host.docker.internal:5432/sardoba_bot`

Ishga tushirish:
```bash
docker compose up -d --build
```

Log ko'rish:
```bash
docker compose logs -f bot
```

To'xtatish:
```bash
docker compose down
```

### Local test uchun
Ichki Postgres va `pgAdmin` bilan ko'tarish:
```bash
docker compose -f docker-compose.yml -f docker-compose.local.yml up -d --build
```

`pgAdmin` kerak bo'lsa:
```bash
docker compose -f docker-compose.yml -f docker-compose.local.yml --profile tools up -d
```

This is a comprehensive Telegram bot for managing restaurant operations with two user roles: Admin and Cashier.

## Features

### For Cashiers:
- Shift opening and closing
- Location selection (4 branches)
- Workplace status reporting with photos
- Daily reporting (sales, debts, expenses, card payments, etc.)
- Shift reconciliation process

### For Admins:
- Manage all cashiers
- Monitor shift openings/closings
- Generate reports
- Approve cashier requests
- Modify information
- Export data to Excel/PDF

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd bot1
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Set up the database (PostgreSQL):
```bash
psql -U sardoba -d sardoba_bot -f sql/postgres/schema.sql
```

4. Configure environment variables:
```bash
cp .env.example .env
# Edit .env with your Telegram bot token and database credentials
```

5. Run the bot:
```bash
python -m sardoba_bot
```

## Configuration

Create a `.env` file with the following variables:

```env
# Telegram Bot Configuration
TELEGRAM_BOT_TOKEN=YOUR_BOT_TOKEN_HERE

# Production / Docker
DATABASE_URL=postgresql://sardoba:sardoba@host.docker.internal:5432/sardoba_bot

# Fallback database configuration
DB_HOST=localhost
DB_NAME=sardoba_bot
DB_USER=sardoba
DB_PASSWORD=sardoba
DB_PORT=5432
```

## Database Schema

The bot uses the following tables:
- `users` - stores admin and cashier information
- `locations` - restaurant branch locations
- `shifts` - tracks cashier shifts
- `reports` - daily reports and financial data
- `images` - uploaded images for reports
- `approval_requests` - cashier registration requests
- `bot_settings` - bot yuboradigan Telegram guruh sozlamalari

## Usage

1. Start the bot with `/start`
2. Select role (Admin/Cashier)
3. Register with personal information and password
4. For cashiers, wait for admin approval
5. Use the appropriate menu based on your role

### For Cashiers:
- Open shift with opening amount
- Upload required photos of workplace status
- Submit daily reports (sales, debts, expenses, etc.)
- Close shift when finished

### For Admins:
- View all cashiers
- Monitor shift activity
- Review and approve cashier requests
- Generate and export reports
- Botni kerakli guruhga ulash uchun botni guruhga qo'shib, o'sha guruh ichida `/setgroup` yuboring

## File Structure

```
sardoba_kassa_bot/
├── sardoba_bot/
│   ├── telegram/      # Telegram bot logic and entrypoint
│   ├── db/            # Database connection and SQL query constants
│   ├── services/      # Export and reporting services
│   ├── common/        # Shared utility helpers
│   └── core/          # Conversation states and menu constants
├── sql/
│   ├── postgres/      # PostgreSQL schema
│   └── mysql_legacy/  # Legacy MySQL schema
├── tests/             # Automated tests
├── docker/            # Container startup scripts
├── bot.py             # Backward-compatible wrapper
├── db_config.py       # Backward-compatible wrapper
├── export_utils.py    # Backward-compatible wrapper
├── utils.py           # Backward-compatible wrapper
└── README.md          # This file
```

## Technologies Used

- Python 3.x
- python-telegram-bot
- PostgreSQL
- openpyxl (for Excel exports)
- Pillow (for Excel image exports)
- reportlab (for PDF generation)

## Security Features

- Password hashing with salt
- Input validation
- Role-based access control
- Secure database connections

## License

This project is licensed under the MIT License.
