from flask import Flask, render_template
from subprocess import Popen, PIPE

app = Flask(__name__)

@app.route('/run_script')
def run_script():
    process = Popen(['python', 'path/to/your/script.py'], stdout=PIPE, stderr=PIPE)
    stdout, stderr = process.communicate()
    return stdout

if __name__ == '__main__':
    app.run(debug=True)
