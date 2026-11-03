import os
from flask import Flask, jsonify, render_template_string
from flask_cors import CORS
from supabase import create_client, Client
from postgrest.exceptions import APIError
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
CORS(app)

# Validate environment variables
SUPABASE_URL = os.environ.get("SUPABASE_URL")
SUPABASE_KEY = os.environ.get("SUPABASE_KEY")

if not SUPABASE_URL or not SUPABASE_KEY:
    raise ValueError(
        "Missing Supabase credentials. Please ensure SUPABASE_URL and SUPABASE_KEY are set in your .env file."
    )

supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Todos - Supabase App</title>
    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; max-width: 600px; margin: 40px auto; padding: 0 20px; color: #333; line-height: 1.5; }
        h1 { color: #1e3a8a; border-bottom: 2px solid #e2e8f0; padding-bottom: 10px; }
        ul { list-style-type: none; padding: 0; }
        li { background: #f8fafc; border: 1px solid #e2e8f0; margin-bottom: 8px; padding: 12px 16px; border-radius: 6px; }
        .empty { color: #64748b; font-style: italic; }
        .error { background: #fee2e2; border: 1px solid #f87171; color: #991b1b; padding: 16px; border-radius: 6px; }
        code { background: #e2e8f0; padding: 2px 6px; border-radius: 4px; font-size: 0.9em; }
        pre { background: #0f172a; color: #e2e8f0; padding: 14px; border-radius: 6px; overflow-x: auto; font-size: 0.85em; }
    </style>
</head>
<body>
    <h1>Todos</h1>
    {% if error %}
        <div class="error">
            <h3>Database Note</h3>
            <p>{{ error }}</p>
            {% if table_missing %}
            <p>The <code>todos</code> table does not exist in your Supabase database yet. To create it, run this SQL in your Supabase SQL Editor:</p>
            <pre>create table if not exists todos (
  id bigint generated always as identity primary key,
  name text not null,
  is_complete boolean default false,
  created_at timestamptz default now() not null
);

-- Enable Row Level Security (RLS)
alter table todos enable row level security;

-- Policy to allow read access for anon and authenticated users
create policy "Allow read access" on todos
  for select
  to anon, authenticated
  using (true);

-- Policy to allow inserting todos
create policy "Allow insert access" on todos
  for insert
  to anon, authenticated
  with check (true);

-- Sample row
insert into todos (name) values ('Complete Gantt chart setup');</pre>
            {% endif %}
        </div>
    {% elif not todos %}
        <p class="empty">No todos found in the database.</p>
    {% else %}
        <ul>
        {% for todo in todos %}
            <li>{{ todo.name or todo.title or todo.task or 'Untitled' }}</li>
        {% endfor %}
        </ul>
    {% endif %}
</body>
</html>
"""

@app.route('/')
def index():
    try:
        response = supabase.table('todos').select("*").execute()
        todos = response.data or []
        return render_template_string(HTML_TEMPLATE, todos=todos, error=None)
    except APIError as e:
        error_msg = e.message if hasattr(e, 'message') else str(e)
        table_missing = "PGRST205" in str(e) or "schema cache" in error_msg
        return render_template_string(HTML_TEMPLATE, todos=[], error=error_msg, table_missing=table_missing), 200
    except Exception as e:
        return render_template_string(HTML_TEMPLATE, todos=[], error=str(e), table_missing=False), 500

@app.route('/api/todos', methods=['GET'])
def get_todos():
    try:
        response = supabase.table('todos').select("*").execute()
        return jsonify({"status": "success", "data": response.data})
    except APIError as e:
        return jsonify({"status": "error", "message": e.message if hasattr(e, 'message') else str(e)}), 400
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
