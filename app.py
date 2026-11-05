import sys
import os

# Ensure the project root directory is always on Python's search path
PROJECT_DIR = os.path.abspath(os.path.dirname(__file__))
if PROJECT_DIR not in sys.path:
    sys.path.insert(0, PROJECT_DIR)

from flask import Flask
from config import Config
from models import db, User
from flask_login import LoginManager

def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)

    # Initialize Database
    db.init_app(app)

    # Initialize Login Manager
    login_manager = LoginManager()
    login_manager.login_view = 'auth.login'
    login_manager.login_message_category = 'warning'
    login_manager.init_app(app)

    @login_manager.user_loader
    def load_user(user_id):
        return User.query.get(int(user_id))

    # Register Blueprints
    from routes.auth import auth_bp
    from routes.main import main_bp
    from routes.booking import booking_bp 
    from routes.admin import admin_bp

    app.register_blueprint(auth_bp)
    app.register_blueprint(main_bp)
    app.register_blueprint(booking_bp)
    app.register_blueprint(admin_bp)

    # Context processors to inject global template values
    @app.context_processor
    def inject_global_data():
        from datetime import datetime
        import config
        return {
            'current_year': datetime.utcnow().year,
            'admin_name': config.Config.ADMIN_NAME,
            'admin_phone': config.Config.ADMIN_PHONE,
            'admin_email': config.Config.ADMIN_EMAIL,
            'admin_phonepe_qr': config.Config.ADMIN_PHONEPE_QR,
            'registration_fee': config.Config.REGISTRATION_FEE,
            'executive_leadership': config.Config.EXECUTIVE_LEADERSHIP,
        }

    # Automatically create tables and seed database
    with app.app_context():
        db.create_all()
        from utils.db_seeder import seed_database
        seed_database()

    return app

app = create_app()

if __name__ == '__main__':
    # Ensure static upload folder exists
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    app.run(debug=True, port=5000)
