import json
import requests

def sync_to_google_sheet(webhook_url, students, school_name="វិទ្យាល័យ ហ៊ុន សែន អង្គប្រីយ៍"):
    """
    Sends all student records to Google Sheet via Google Apps Script Webhook.
    """
    if not webhook_url or not webhook_url.startswith("http"):
        return {"success": False, "message": "សូមបញ្ចូល Google Apps Script Webhook URL ជាមុនសិន"}

    payload = {
        "action": "sync_all",
        "school": school_name,
        "total": len(students),
        "students": students
    }

    try:
        # Note: Google Apps Script Web Apps redirect on POST (302 redirect), so allow_redirects=True is required
        headers = {'Content-Type': 'application/json'}
        resp = requests.post(webhook_url, json=payload, headers=headers, timeout=15, allow_redirects=True)
        
        try:
            res_json = resp.json()
            return {"success": True, "message": res_json.get("message", "បានធ្វើសមកាលកម្មទិន្នន័យទៅ Google Sheet ដោយជោគជ័យ!")}
        except Exception:
            if resp.status_code in [200, 302]:
                return {"success": True, "message": f"បានបញ្ជូនទិន្នន័យសិស្សចំនួន {len(students)} នាក់ទៅ Google Sheet រួចរាល់!"}
            return {"success": False, "message": f"Server ឆ្លើយតបកូដ: {resp.status_code}"}
    except requests.exceptions.Timeout:
        return {"success": False, "message": "ការតភ្ជាប់ទៅកាន់ Google Sheet ផុតកំណត់ពេល (Timeout)។ សូមពិនិត្យមើលអ៊ីនធឺណិត។"}
    except Exception as e:
        return {"success": False, "message": f"បរាជ័យក្នុងការតភ្ជាប់៖ {str(e)}"}

def pull_from_google_sheet(webhook_url):
    """
    Fetches records from Google Sheet via Google Apps Script Webhook.
    """
    if not webhook_url or not webhook_url.startswith("http"):
        return {"success": False, "message": "សូមបញ្ចូល Google Apps Script Webhook URL ជាមុនសិន", "students": []}

    try:
        resp = requests.get(webhook_url, params={"action": "get_all"}, timeout=15, allow_redirects=True)
        res_json = resp.json()
        if res_json.get("status") == "success" or "students" in res_json:
            return {"success": True, "students": res_json.get("students", []), "message": "ទាញទិន្នន័យពី Google Sheet ជោគជ័យ!"}
        return {"success": False, "message": res_json.get("message", "មិនអាចទាញទិន្នន័យបានទេ"), "students": []}
    except Exception as e:
        return {"success": False, "message": f"បរាជ័យ៖ {str(e)}", "students": []}
