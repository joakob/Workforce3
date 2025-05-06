from flask import Blueprint, render_template

layout_bp = Blueprint('layout', __name__)

@layout_bp.route('/layout')
def layout():
    return render_template('layout.html', title='Layout')