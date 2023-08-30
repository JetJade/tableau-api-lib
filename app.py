from flask import Flask, render_template
app = Flask(__name__)
import main_2

tableau_ext = main_2.TableauExtension()
in_execution = 0

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/execute_function')
def execute_function():
    global in_execution
    in_execution = 1
    global tableau_ext
    tableau_ext = main_2.TableauExtension()
    result = tableau_ext.create_pdf()
    in_execution = 0
    return result

@app.route('/get_status')
def get_status():
    if in_execution == 1:
        return str(tableau_ext.check_status())
    else:
        return str(0)

if __name__ == '__main__':
    app.run(debug=True)