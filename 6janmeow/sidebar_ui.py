from flask import render_template_string

class SidebarUI:
    @staticmethod
    def get_sidebar():
        """Generate sidebar HTML."""
        return """
        <div class="sidebar">
            <h2>Highlighter Tool</h2>
            <ul>
                <li><a href="/">Home</a></li>
                <li><a href="/upload">Upload File</a></li>
                <li><a href="/results">View Results</a></li>
                <li><a href="/files">Output Files</a></li>
                <li><a href="/run">Run Analysis</a></li>
            </ul>
            
            <div class="sidebar-section">
                <h3>Output Stats</h3>
                <div id="stats">
                    <p>Loading stats...</p>
                </div>
            </div>
        </div>
        """
    
    @staticmethod
    def get_main_content(content):
        """Wrap content with sidebar layout."""
        return f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>Survey Highlighter Tool</title>
            <style>
                body {{ 
                    margin: 0; 
                    font-family: Arial, sans-serif;
                    display: flex;
                }}
                .sidebar {{
                    width: 250px;
                    background: #2c3e50;
                    color: white;
                    padding: 20px;
                    height: 100vh;
                    position: fixed;
                    overflow-y: auto;
                }}
                .sidebar h2 {{
                    color: #ecf0f1;
                    border-bottom: 2px solid #3498db;
                    padding-bottom: 10px;
                }}
                .sidebar ul {{
                    list-style: none;
                    padding: 0;
                }}
                .sidebar li {{
                    margin: 15px 0;
                }}
                .sidebar a {{
                    color: #bdc3c7;
                    text-decoration: none;
                    font-size: 16px;
                    display: block;
                    padding: 10px;
                    border-radius: 5px;
                    transition: all 0.3s;
                }}
                .sidebar a:hover {{
                    background: #34495e;
                    color: white;
                }}
                .main-content {{
                    margin-left: 270px;
                    padding: 30px;
                    flex: 1;
                }}
                .card {{
                    background: white;
                    border-radius: 10px;
                    padding: 20px;
                    margin: 20px 0;
                    box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                }}
                .btn {{
                    background: #3498db;
                    color: white;
                    padding: 10px 20px;
                    border: none;
                    border-radius: 5px;
                    cursor: pointer;
                    text-decoration: none;
                    display: inline-block;
                }}
                .btn:hover {{
                    background: #2980b9;
                }}
                .file-list {{
                    max-height: 400px;
                    overflow-y: auto;
                    border: 1px solid #ddd;
                    padding: 10px;
                    border-radius: 5px;
                }}
                table {{
                    width: 100%;
                    border-collapse: collapse;
                    margin: 20px 0;
                }}
                th, td {{
                    border: 1px solid #ddd;
                    padding: 10px;
                    text-align: left;
                }}
                th {{
                    background: #f2f2f2;
                }}
                .highlight-cell {{
                    background-color: #FBE2D5 !important;
                }}
                .green-highlight-cell {{
                    background-color: #DAF2D0 !important;
                }}
            </style>
        </head>
        <body>
            {SidebarUI.get_sidebar()}
            <div class="main-content">
                {content}
            </div>
            
            <script>
                // Load stats dynamically
                fetch('/api/stats')
                    .then(response => response.json())
                    .then(data => {{
                        document.getElementById('stats').innerHTML = `
                            <p>Total Files: ${{data.total_files}}</p>
                            <p>Highlighted Files: ${{data.highlighted_files}}</p>
                            <p>Green Files: ${{data.green_files}}</p>
                            <p>CSV Files: ${{data.csv_files}}</p>
                        `;
                    }});
            </script>
        </body>
        </html>
        """