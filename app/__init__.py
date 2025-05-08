from flask import Flask

def create_app():
    app = Flask(__name__)
    
    from app.static.home import home_bp

    app.register_blueprint(home_bp)

    return app