# SemaineHors-Cadre - Workshop Planning Optimizer

A web application designed to optimize workshop schedules for schools, matching students to activities based on their preferences while respecting capacity constraints.

## Overview

This application allows educational institutions to:

1. Upload workshop and student preference data via a pre-formatted Excel template
2. Run an advanced optimization algorithm that assigns students to workshops
3. Automatically generate a complete schedule that maximizes student satisfaction
4. Download the results in Excel format with detailed statistics

The optimizer uses linear programming techniques to create the best possible schedule while respecting various constraints like:
- Maximum capacity of each workshop
- Student preferences (including vetoes)
- Session timing conflicts
- Workshop duration requirements

## Features

- **User-friendly web interface**: Simple upload/download functionality with real-time feedback
- **Intelligent scheduling**: Maximizes student preference satisfaction while balancing workshop attendance
- **Comprehensive output**: Generates multiple Excel worksheets showing:
  - Student schedules (what each student should attend)
  - Workshop rosters (lists of students assigned to each workshop)
  - Statistical analysis of preference satisfaction
- **Visual feedback**: Charts showing distribution of preference satisfaction
- **Configurable parameters**: Adjustable rewards/penalties to fine-tune optimization behavior

## Project Structure

```
SemaineHors-Cadre/
├── Dockerfile                # Docker configuration for deployment
├── planning_final.xlsx       # Sample output file
├── preferences.ipynb         # Jupyter notebook for analysis
├── requirements.txt          # Project-level dependencies
├── solver.py                 # Standalone solver script
├── Template.xlsx             # Template file for data input
└── webapp/                   # Web application folder
    ├── app.py                # Flask application main file
    ├── requirements.txt      # Web application dependencies
    ├── solver_logic.py       # Optimization algorithm implementation
    ├── static/               # Static files (templates, etc.)
    ├── templates/            # HTML templates
    │   └── index.html        # Main web interface
    ├── uploads/              # Temporary folder for uploaded files
    └── results/              # Output files storage
```

## Prerequisites

- Python 3.10+
- Docker (for containerized deployment)

## Local Development Setup

1. Clone the repository:
   ```
   git clone <repository-url>
   cd SemaineHors-Cadre
   ```

2. Create and activate a virtual environment:
   ```
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install dependencies:
   ```
   cd webapp
   pip install -r requirements.txt
   ```

4. Run the Flask development server:
   ```
   python app.py
   ```

5. Access the application at http://localhost:5000

## Docker Deployment

### Building and Deploying in One Command

```bash
docker build -t semaine-hors-cadre:latest . && docker run -d --name semaine-hors-cadre -p 80:5000 -v $(pwd)/webapp/results:/app/results --restart unless-stopped semaine-hors-cadre:latest
```

### Step-by-Step Deployment

1. Build the Docker image:
   ```
   docker build -t semaine-hors-cadre:latest .
   ```

2. Run the container:
   ```
   docker run -d --name semaine-hors-cadre -p 80:5000 -v $(pwd)/webapp/results:/app/results --restart unless-stopped semaine-hors-cadre:latest
   ```

   This command:
   - Runs the container in detached mode (`-d`)
   - Names it "semaine-hors-cadre"
   - Maps port 80 on your server to port 5000 in the container
   - Creates a volume to persist generated results
   - Sets the container to restart automatically unless manually stopped

3. Access the application at http://your-server-ip

## How to Use

1. **Download Template**: Get the Excel template by clicking "Télécharger Template.xlsx"

2. **Fill the Template**:
   - In the "Ateliers" sheet: Enter workshop details (code, description, teacher, room, capacity)
   - In the "Preferences" sheet: Enter student information and their preferences (1 for preference, -1 for veto)

3. **Upload and Process**: Upload your completed template file and click "Lancer l'Optimisation"

4. **View Results**: Once processing is complete, you'll see statistics and a chart showing preference distribution

5. **Download Results**: Click "Télécharger le Planning Complet (.xlsx)" to get the full schedule

## Technology Stack

- **Backend**: Python, Flask
- **Optimization**: PuLP (linear programming)
- **Data Processing**: Pandas
- **Frontend**: HTML, CSS, JavaScript
- **Styling**: Bootstrap 5
- **Visualization**: Chart.js
- **Containerization**: Docker

## Security Considerations

For production deployment:
- Set a secure `FLASK_SECRET_KEY` environment variable
- Consider adding authentication if used in a sensitive environment
- Use HTTPS with a proper SSL certificate
- Consider adding nginx as a reverse proxy for improved security and performance

## License

[MIT License](LICENSE)

## Contributing

Contributions welcome! Please feel free to submit a Pull Request.