from flask import Flask, request, send_from_directory, send_file
from werkzeug.utils import secure_filename
import json
import os
from flask_cors import CORS

# import doc-assist
from doc_assist_be.doc_assist import application as doc_assist_app

application = Flask(__name__)
cors = CORS(application)
application.register_blueprint(doc_assist_app, url_prefix="/doc-assist")


@application.route("/doc-assist/<mount_type>")
@application.route("/doc-assist")
def index(mount_type=None):
    application.static_folder = "build/static"
    return send_from_directory(
        os.path.join(application.root_path, "build"), "index.html"
    )


if __name__ == "__main__":
    application.run(host="0.0.0.0")
