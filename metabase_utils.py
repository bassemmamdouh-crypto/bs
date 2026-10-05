from io import StringIO
import json

import pandas as pd
import requests


# NOTE:
# Credentials are intentionally hard-coded here per request.
MARBAH_BASE_URL = "https://bi.marbah.info/api"
MARBAH_USERNAME = "marbah_username"
MARBAH_PASSWORD = "marbah_password"

AFRICA_BASE_URL = "https://bi.maxabma.com/api"
AFRICA_USERNAME = "africa_username"
AFRICA_PASSWORD = "africa_password"


def ret_metabase(country, question, use_query=False, filters=None, warehouse=None):
    """
    Retrieve data from Metabase by executing a question or query.

    Supports:
    - marbah / egypt Metabase instance
    - default Africa Metabase instance
    """
    # Keep this call if the broader environment expects it.
    initialize_env()

    question_id = str(question)
    filters = filters or {}
    country_lower = str(country).lower()

    if country_lower in {"marbah", "egypt"}:
        base_url = MARBAH_BASE_URL
        username = MARBAH_USERNAME
        password = MARBAH_PASSWORD
    else:
        base_url = AFRICA_BASE_URL
        username = AFRICA_USERNAME
        password = AFRICA_PASSWORD

    base_headers = {"Content-Type": "application/json"}

    try:
        s_response = requests.post(
            f"{base_url}/session",
            data=json.dumps({"username": username, "password": password}),
            headers=base_headers,
        )
        s_response.raise_for_status()

        session_token = s_response.json()["id"]
        base_headers["X-Metabase-Session"] = session_token

        params = []
        for name, value in filters.items():
            filter_type, filter_value = value
            param = {"target": ["variable", ["template-tag", name]], "value": filter_value}

            filter_type_lower = filter_type.lower()
            if filter_type_lower == "date":
                param["type"] = "date/range" if isinstance(filter_value, list) else "date/single"
            elif filter_type_lower == "category":
                param["type"] = "category"
            elif filter_type_lower == "text":
                param["type"] = "text"
            elif filter_type_lower == "number":
                param["type"] = "number"
            elif filter_type_lower == "field list":
                param["type"] = "id"
                param["target"] = ["dimension", ["template-tag", name]]

            params.append(param)

        if use_query:
            p_response = requests.get(f"{base_url}/card/{question_id}", headers=base_headers)
            p_response.raise_for_status()
            rj = p_response.json()
            card = rj.get("dataset_query", {})
            query = card.get("native", {}).get("query", "").replace("\n", " ")
            return snowflake_query(country, query, warehouse)

        p_response = requests.post(
            f"{base_url}/card/{question_id}/query/csv",
            json={"parameters": params},
            headers=base_headers,
        )
        p_response.raise_for_status()

        csv_text = p_response.content.decode("utf-8")
        return pd.read_csv(StringIO(csv_text))

    except Exception as e:
        logger.error(f"An error occurred: {e}", exc_info=True)
        raise
