from flask import Flask, send_file, request, jsonify
import yaml

app = Flask(__name__)


@app.route('/')
def index():
    return '''
    <html>
    <head>
        <style>
            body {
                background-color: white;
                color: red;
                text-align: center;
            }

            h1 {
                color: black;
            }

            #setTimeFrequencyButton {
                background-color: black;
                color: red;
                padding: 10px 20px;
                border: none;
                cursor: pointer;
            }

            #updateDatasetButton {
                background-color: black;
                color: red;
                padding: 10px 20px;
                border: none;
                cursor: pointer;
            }

        </style>
        <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
        <script>
            $(document).ready(function() {
                $("#setTimeFrequencyButton").click(function(e) {
                    e.preventDefault();
                    var timeFrequency = $("#timeFrequencyInput").val();
                    $.ajax({
                        type: "POST",
                        url: "/set_time_frequency",
                        data: { time_frequency: timeFrequency },
                        success: function(data) {
                            $("#result").text(data);
                        },
                        error: function(error) {
                            $("#result").text("Error: " + error.responseText);
                        }
                    });
                });

                $("#updateDatasetButton").click(function(e) {
                    e.preventDefault();
                    $.ajax({
                        type: "POST",
                        url: "/update_dataset",
                        success: function(data) {
                            // Display the message after 15 seconds
                            setTimeout(function() {
                                $("#datasetUpdateMessage").text("The internships dataset was updated");
                            }, 15000); // 15,000 milliseconds = 15 seconds
                        },
                        error: function(error) {
                            $("#datasetUpdateMessage").text("Error: " + error.responseText);
                        }
                    });
                });
            });
        </script>
    </head>
    <body>
        <h1>Download Internships Excel File</h1>
        <p><a href="/download">Download Excel</a></p>

        <h2>Set Time Frequency</h2>
        <input type="number" id="timeFrequencyInput" placeholder="Enter time frequency">
        <button id="setTimeFrequencyButton">Set Time Frequency</button>
        <div id="result"></div>

        <h2>Update Dataset</h2>
        <button id="updateDatasetButton">Dataset Update</button>
        <div id="datasetUpdateMessage"></div>
    </body>
    </html>
    '''


@app.route('/download')
def download_excel():
    excel_file_path = 'Stages_DataSet.xlsx'  # Provide the correct path to your Excel file
    return send_file(excel_file_path, as_attachment=True)


@app.route('/set_time_frequency', methods=['POST'])
def set_time_frequency():
    time_frequency = request.form.get('time_frequency')
    if time_frequency is not None and time_frequency.isdigit():
        time_frequency = int(time_frequency)
        if time_frequency > 1000:
            # Load existing credentials from credentials.yml
            with open('credentials.yml', 'r') as file:
                credentials = yaml.safe_load(file)

            # Update the time_frequency in the credentials dictionary
            credentials['time_frequency'] = time_frequency

            # Save the updated credentials dictionary to credentials.yml
            with open('credentials.yml', 'w') as file:
                yaml.dump(credentials, file, default_flow_style=False)

            return f'Time frequency set to {time_frequency} and saved to credentials.yml'
        else:
            return 'Time frequency must be greater than 1000.', 400
    else:
        return 'Invalid input. Please enter a valid integer.', 400


if __name__ == '__main__':
    app.run(debug=True)
