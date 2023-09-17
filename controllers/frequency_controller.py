from flask import Blueprint, request, jsonify

frequency_controller = Blueprint('frequency_controller', __name__)

@frequency_controller.route('/set_frequency', methods=['POST'])
def set_time_frequency():
    # Handle setting time frequency here
    return jsonify({'message': 'Time frequency updated successfully'})
