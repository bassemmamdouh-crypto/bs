from io import StringIO
import logging

import pandas as pd
import requests


logger = logging.getLogger(__name__)

# Iraq Metabase configuration (hardcoded as requested)
IRAQ_METABASE_HOST = "https://bi.marbah.info"
IRAQ_METABASE_USERNAME = "YOUR_IRAQ_METABASE_USERNAME"
IRAQ_METABASE_PASSWORD = "YOUR_IRAQ_METABASE_PASSWORD"


def _safe_response_json(response, context):
    """
    Parse JSON safely and raise a clear error if the response is not JSON.
    """
    try:
        return response.json()
    except ValueError as exc:
        content_type = response.headers.get("Content-Type", "unknown")
        body_preview = response.text[:300].replace("\n", " ").strip()
        raise RuntimeError(
            f"{context} returned non-JSON response "
            f"(status={response.status_code}, content_type={content_type}). "
            f"Body preview: {body_preview}"
        ) from exc


def _session_with_fallback_paths(base_headers):
    """
    Authenticate against Iraq Metabase trying common API path variants.
    """
    payload = {
        "username": IRAQ_METABASE_USERNAME,
        "password": IRAQ_METABASE_PASSWORD,
    }

    # Common reverse-proxy prefixes for Metabase API.
    candidate_api_bases = (
        f"{IRAQ_METABASE_HOST}/api",
        f"{IRAQ_METABASE_HOST}/metabase/api",
        f"{IRAQ_METABASE_HOST}/analytics/api",
        f"{IRAQ_METABASE_HOST}/bi/api",
        f"{IRAQ_METABASE_HOST}/mb/api",
    )

    attempt_errors = []
    for api_base_url in candidate_api_bases:
        session_url = f"{api_base_url}/session"
        try:
            response = requests.post(
                session_url,
                json=payload,
                headers=base_headers,
                timeout=30,
            )
            response.raise_for_status()
            session_json = _safe_response_json(
                response, f"Metabase session endpoint ({session_url})"
            )
            session_token = session_json.get("id")
            if not session_token:
                raise RuntimeError(
                    f"Metabase session response from {session_url} is missing 'id': {session_json}"
                )
            return api_base_url, session_token
        except Exception as exc:
            content_type = "unknown"
            status_code = "unknown"
            body_preview = ""
            if "response" in locals() and response is not None:
                content_type = response.headers.get("Content-Type", "unknown")
                status_code = response.status_code
                body_preview = response.text[:120].replace("\n", " ").strip()
            attempt_errors.append(
                f"{session_url} -> status={status_code}, content_type={content_type}, "
                f"error={exc}, body_preview={body_preview}"
            )

    attempts_summary = " | ".join(attempt_errors)
    raise RuntimeError(
        "Failed to authenticate with Iraq Metabase across known API paths. "
        f"Host={IRAQ_METABASE_HOST}. Attempts: {attempts_summary}"
    )


def ret_metabase(question, use_query=False, filters=None, warehouse=None):
    """
    Retrieve data from Iraq Metabase by executing a question or query.

    Args:
        question (str): The question ID or query to execute in Metabase.
        use_query (bool, optional): If True, execute the underlying native SQL in Snowflake.
        filters (dict, optional): Question filters in the format:
            {'filter_name': ('filter_type', filter_value)}
            Supported types: date, category, text, number, field list.
        warehouse (str, optional): Snowflake warehouse to use when use_query=True.

    Returns:
        pandas.DataFrame: A DataFrame containing the query results.
    """
    if filters is None:
        filters = {}

    question_id = str(question)
    base_headers = {"Content-Type": "application/json"}

    try:
        api_base_url, session_token = _session_with_fallback_paths(base_headers)
        base_headers["X-Metabase-Session"] = session_token

        params = []
        for name, value in filters.items():
            filter_type, filter_value = value
            param = {"target": ["variable", ["template-tag", name]], "value": filter_value}

            filter_type = filter_type.lower()
            if filter_type == "date":
                param["type"] = "date/range" if isinstance(filter_value, list) else "date/single"
            elif filter_type == "category":
                param["type"] = "category"
            elif filter_type == "text":
                param["type"] = "text"
            elif filter_type == "number":
                param["type"] = "number"
            elif filter_type == "field list":
                param["type"] = "id"
                param["target"] = ["dimension", ["template-tag", name]]

            params.append(param)

        if use_query:
            p_response = requests.get(
                f"{api_base_url}/card/{question_id}",
                headers=base_headers,
            )
            p_response.raise_for_status()
            card_response = _safe_response_json(p_response, "Metabase card endpoint")
            card = card_response.get("dataset_query", {})
            query = card.get("native", {}).get("query", "").replace("\n", " ")
            if not query:
                raise RuntimeError(
                    "No native SQL query was found in Metabase card dataset_query."
                )
            return snowflake_query("iraq", query, warehouse)

        p_response = requests.post(
            f"{api_base_url}/card/{question_id}/query/csv",
            json={"parameters": params},
            headers=base_headers,
        )
        p_response.raise_for_status()

        csv_buffer = StringIO(p_response.content.decode("utf-8"))
        return pd.read_csv(csv_buffer)

    except Exception as e:
        logger.error(f"An error occurred: {e}", exc_info=True)
        raise
