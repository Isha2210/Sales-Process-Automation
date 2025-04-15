# Enhanced Flask server for tracking email opens and clicks
from flask import Flask, request, redirect, send_file, jsonify, abort
import json
import os
import re
import urllib.parse
from datetime import datetime
import logging
from logging.handlers import RotatingFileHandler
import threading

app = Flask(__name__)

# Set up enhanced logging with rotation
log_file = 'tracking_server.log'
log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

# File handler with rotation (10MB max size, keep 5 backup files)
file_handler = RotatingFileHandler(log_file, maxBytes=10 * 1024 * 1024, backupCount=5)
file_handler.setFormatter(log_formatter)
file_handler.setLevel(logging.INFO)

# Console handler
console_handler = logging.StreamHandler()
console_handler.setFormatter(log_formatter)
console_handler.setLevel(logging.INFO)

# Root logger configuration
logging.basicConfig(
    level=logging.INFO,
    handlers=[file_handler, console_handler]
)
logger = logging.getLogger(__name__)

# Lock for thread-safe file operations
file_lock = threading.Lock()

# Configuration
TRACKING_DATA_DIR = 'tracking_data'
DEFAULT_REDIRECT_URL = 'https://www.example.com'
PIXEL_FILE = 'pixel.png'
VALID_TRACKING_ID_PATTERN = re.compile(r'^[a-zA-Z0-9]+_[a-zA-Z0-9]+$')

# Ensure data directory exists
os.makedirs(TRACKING_DATA_DIR, exist_ok=True)


def is_valid_tracking_id(tracking_id):
    """Validate tracking ID format"""
    return bool(VALID_TRACKING_ID_PATTERN.match(tracking_id))


def is_valid_url(url):
    """Basic URL validation"""
    try:
        result = urllib.parse.urlparse(url)
        return all([result.scheme in ['http', 'https'], result.netloc])
    except:
        return False


@app.route('/')
def index():
    """Home page / dashboard"""
    return jsonify({
        "status": "running",
        "server": "Email Tracking Server",
        "version": "1.0",
        "available_endpoints": [
            "/health - Server health check",
            "/track/pixel/<tracking_id> - Track email opens",
            "/track/click/<tracking_id>?url=... - Track link clicks",
            "/stats/<campaign_id> - View campaign statistics"
        ]
    })


@app.route('/health')
def health_check():
    """Simple health check endpoint"""
    return jsonify({"status": "healthy", "timestamp": datetime.now().isoformat()})


@app.route('/track/pixel/<tracking_id>')
def track_open(tracking_id):
    """Track when an email is opened"""
    # Validate tracking ID
    if not is_valid_tracking_id(tracking_id):
        logger.warning(f"Invalid tracking ID format: {tracking_id}")
        abort(400)

    # Get client IP and user agent for enhanced tracking
    client_ip = request.remote_addr
    user_agent = request.headers.get('User-Agent', 'Unknown')

    logger.info(f"Email opened: {tracking_id} | IP: {client_ip} | Agent: {user_agent}")

    # Update tracking data
    update_tracking_data(tracking_id, 'opened', client_ip, user_agent)

    # Return a 1x1 transparent pixel
    try:
        return send_file(PIXEL_FILE, mimetype='image/png')
    except FileNotFoundError:
        logger.error(f"Tracking pixel not found: {PIXEL_FILE}")
        abort(500)


@app.route('/track/click/<tracking_id>')
def track_click(tracking_id):
    """Track when a link is clicked and redirect"""
    # Validate tracking ID
    if not is_valid_tracking_id(tracking_id):
        logger.warning(f"Invalid tracking ID format: {tracking_id}")
        abort(400)

    # Get redirect URL with fallback
    url = request.args.get('url', DEFAULT_REDIRECT_URL)

    # Validate URL
    if not is_valid_url(url):
        logger.warning(f"Invalid URL in click request: {url}")
        url = DEFAULT_REDIRECT_URL

    # Get client IP and user agent for enhanced tracking
    client_ip = request.remote_addr
    user_agent = request.headers.get('User-Agent', 'Unknown')

    logger.info(f"Link clicked: {tracking_id} | Redirect: {url} | IP: {client_ip} | Agent: {user_agent}")

    # Update tracking data
    update_tracking_data(tracking_id, 'clicked', client_ip, user_agent, url)

    # Redirect to the original URL
    return redirect(url)


@app.route('/stats/<campaign_id>')
def campaign_stats(campaign_id):
    """API endpoint to fetch campaign statistics"""
    # Basic validation
    if not re.match(r'^[a-zA-Z0-9]+$', campaign_id):
        abort(400)

    data_file = os.path.join(TRACKING_DATA_DIR, f'campaign_data_{campaign_id}.json')

    try:
        with file_lock:
            if not os.path.exists(data_file):
                return jsonify({"error": "Campaign not found"}), 404

            with open(data_file, 'r') as f:
                tracking_data = json.load(f)

        # Calculate statistics
        total_recipients = len(tracking_data)
        total_opens = sum(1 for data in tracking_data.values() if data.get('opened', False))
        total_clicks = sum(1 for data in tracking_data.values() if data.get('clicked', False))

        stats = {
            "campaign_id": campaign_id,
            "total_recipients": total_recipients,
            "total_opens": total_opens,
            "total_clicks": total_clicks,
            "open_rate": round(total_opens / total_recipients * 100, 2) if total_recipients > 0 else 0,
            "click_rate": round(total_clicks / total_recipients * 100, 2) if total_recipients > 0 else 0,
            "click_to_open_rate": round(total_clicks / total_opens * 100, 2) if total_opens > 0 else 0,
        }

        return jsonify(stats)

    except Exception as e:
        logger.error(f"Error generating stats for campaign {campaign_id}: {e}")
        return jsonify({"error": "Internal server error"}), 500


def update_tracking_data(tracking_id, action, ip=None, user_agent=None, url=None):
    """Update the tracking data file with the new event"""
    # Parse the campaign ID from the tracking ID
    try:
        campaign_id = tracking_id.split('_')[0]
    except IndexError:
        logger.error(f"Invalid tracking ID format: {tracking_id}")
        return

    data_file = os.path.join(TRACKING_DATA_DIR, f'campaign_data_{campaign_id}.json')

    try:
        # Use threading lock to prevent race conditions
        with file_lock:
            # Create file with empty data if it doesn't exist
            if not os.path.exists(data_file):
                with open(data_file, 'w') as f:
                    json.dump({}, f)

            # Load the existing data
            with open(data_file, 'r') as f:
                tracking_data = json.load(f)

            # Get current timestamp
            current_time = datetime.now().isoformat()

            # Create entry if it doesn't exist
            if tracking_id not in tracking_data:
                tracking_data[tracking_id] = {
                    "first_seen": current_time,
                    "events": []
                }

            # Add the new event
            event_data = {
                "action": action,
                "timestamp": current_time,
                "ip": ip,
                "user_agent": user_agent
            }

            # Add URL for click events
            if action == 'clicked' and url:
                event_data["url"] = url

            # Update tracking record
            tracking_data[tracking_id]["events"].append(event_data)
            tracking_data[tracking_id][action] = True
            tracking_data[tracking_id]["last_activity"] = current_time

            # Save the updated data
            with open(data_file, 'w') as f:
                json.dump(tracking_data, f, indent=4)

            logger.info(f"Updated tracking data for {tracking_id}: {action}")

    except Exception as e:
        logger.error(f"Error updating tracking data for {tracking_id}: {e}")


# Create the 1x1 transparent pixel
def create_tracking_pixel():
    """Create a 1x1 transparent pixel for tracking email opens"""
    try:
        # Try using PIL if available
        try:
            from PIL import Image
            img = Image.new('RGBA', (1, 1), (0, 0, 0, 0))
            img.save(PIXEL_FILE)
            logger.info(f"Created tracking pixel using PIL: {PIXEL_FILE}")
        except ImportError:
            # Fallback to creating a simple transparent GIF
            with open(PIXEL_FILE, 'wb') as f:
                # Minimal transparent GIF
                f.write(
                    b'\x47\x49\x46\x38\x39\x61\x01\x00\x01\x00\x80\x00\x00\x00\x00\x00\x00\x00\x00\x21\xF9\x04\x01\x00\x00\x00\x00\x2C\x00\x00\x00\x00\x01\x00\x01\x00\x00\x02\x02\x44\x01\x00\x3B')
            logger.info(f"Created tracking pixel using raw bytes: {PIXEL_FILE}")
    except Exception as e:
        logger.error(f"Failed to create tracking pixel: {e}")


@app.errorhandler(404)
def page_not_found(e):
    return jsonify({"error": "Not found"}), 404


@app.errorhandler(400)
def bad_request(e):
    return jsonify({"error": "Bad request"}), 400


@app.errorhandler(500)
def server_error(e):
    return jsonify({"error": "Internal server error"}), 500


if __name__ == '__main__':
    # Create tracking pixel if it doesn't exist
    if not os.path.exists(PIXEL_FILE):
        create_tracking_pixel()

    app.run(debug=True, host='0.0.0.0', port=5000)