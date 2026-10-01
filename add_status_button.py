from pathlib import Path

p = Path("templates/worker.html")
text = p.read_text(encoding="utf-8")

old = """                            <a
                                href="{{ url_for('attendance', worker_id=w['id']) }}"
                                class="small-btn"
                                style="background:#dbeafe;color:#1e40af;"
                            >
                                {{ tr('Attendance') }}
                            </a>"""

new = """                            <a
                                href="{{ url_for('attendance', worker_id=w['id']) }}"
                                class="small-btn"
                                style="background:#dbeafe;color:#1e40af;"
                            >
                                {{ tr('Attendance') }}
                            </a>

                            <form
                                method="post"
                                action="{{ url_for('toggle_worker_status', worker_id=w['id']) }}"
                                style="display:inline;"
                            >
                                <button
                                    type="submit"
                                    class="small-btn"
                                    style="{% if (w['status'] or 'Active') == 'Active' %}background:#fee2e2;color:#991b1b;{% else %}background:#dcfce7;color:#166534;{% endif %}"
                                >
                                    {% if (w['status'] or 'Active') == 'Active' %}
                                        Inactive
                                    {% else %}
                                        Active
                                    {% endif %}
                                </button>
                            </form>"""

if old not in text:
    print("ERROR: Expected Action block পাওয়া যায়নি। কোনো পরিবর্তন করা হয়নি।")
else:
    text = text.replace(old, new, 1)
    p.write_text(text, encoding="utf-8")
    print("STATUS BUTTON ADDED SUCCESSFULLY")
