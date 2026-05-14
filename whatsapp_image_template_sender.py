import argparse
import json
import time
from pathlib import Path

import pandas as pd
import requests


GRAPH_VERSION = "v22.0"


def parse_args():
    parser = argparse.ArgumentParser(
        description=(
            "Upload an image to WhatsApp Cloud API, then send a template message "
            "with image header to each phone number in an Excel file."
        )
    )
    parser.add_argument("--phone-number-id", required=True, help="WhatsApp Phone Number ID")
    parser.add_argument("--access-token", required=True, help="Meta access token")
    parser.add_argument("--template-name", required=True, help="Template name from Meta")
    parser.add_argument("--excel-path", required=True, help="Path to source Excel file")
    parser.add_argument(
        "--phone-column",
        default="NORMALIZED_PHONE",
        help="Column name that contains destination phone numbers",
    )
    parser.add_argument(
        "--sheet-name",
        default=0,
        help="Excel sheet name or index (default: first sheet)",
    )
    parser.add_argument(
        "--language-code",
        default="en",
        help="Template language code (default: en)",
    )
    parser.add_argument(
        "--image-path",
        required=True,
        help="Local image path to upload for template header",
    )
    parser.add_argument(
        "--log-path",
        default="whatsapp_send_logs.xlsx",
        help="Output Excel log path (default: whatsapp_send_logs.xlsx)",
    )
    parser.add_argument(
        "--sleep-seconds",
        type=float,
        default=0.2,
        help="Delay between sends for rate limiting (default: 0.2)",
    )
    return parser.parse_args()


def safe_json(response):
    try:
        return response.json()
    except ValueError:
        return {"raw_text": response.text}


def get_image_mime_type(image_path: Path) -> str:
    ext = image_path.suffix.lower()
    if ext in {".jpg", ".jpeg"}:
        return "image/jpeg"
    if ext == ".png":
        return "image/png"
    raise ValueError("Unsupported image format. Use JPG, JPEG, or PNG.")


def precheck_phone_number_id(phone_number_id: str, access_token: str):
    url = f"https://graph.facebook.com/{GRAPH_VERSION}/{phone_number_id}"
    headers = {"Authorization": f"Bearer {access_token}"}
    params = {"fields": "id,display_phone_number,verified_name"}
    response = requests.get(url, headers=headers, params=params, timeout=20)
    data = safe_json(response)
    if not response.ok:
        raise RuntimeError(
            "Phone number precheck failed. "
            f"status={response.status_code}, response={json.dumps(data, ensure_ascii=True)}"
        )
    print("Precheck status: 200")
    print(f"Precheck body: {json.dumps(data, ensure_ascii=True)}")


def upload_image(phone_number_id: str, access_token: str, image_path: Path) -> str:
    url = f"https://graph.facebook.com/{GRAPH_VERSION}/{phone_number_id}/media"
    headers = {"Authorization": f"Bearer {access_token}"}
    mime_type = get_image_mime_type(image_path)

    with image_path.open("rb") as image_file:
        files = {"file": (image_path.name, image_file, mime_type)}
        form_data = {"messaging_product": "whatsapp", "type": "image"}
        response = requests.post(
            url,
            headers=headers,
            files=files,
            data=form_data,
            timeout=30,
        )

    data = safe_json(response)
    if not response.ok:
        raise RuntimeError(
            "Image upload failed. "
            f"status={response.status_code}, response={json.dumps(data, ensure_ascii=True)}"
        )

    image_id = data.get("id")
    if not image_id:
        raise RuntimeError(
            f"Image upload succeeded but no media id was returned: {json.dumps(data, ensure_ascii=True)}"
        )

    print(f"Image uploaded successfully. media_id={image_id}")
    return image_id


def build_template_payload(
    to_number: str,
    template_name: str,
    language_code: str,
    image_id: str,
):
    return {
        "messaging_product": "whatsapp",
        "to": to_number,
        "type": "template",
        "template": {
            "name": template_name,
            "language": {"code": language_code},
            "components": [
                {
                    "type": "header",
                    "parameters": [
                        {
                            "type": "image",
                            "image": {"id": image_id},
                        }
                    ],
                }
            ],
        },
    }


def send_template_messages(
    contacts_df: pd.DataFrame,
    phone_column: str,
    phone_number_id: str,
    access_token: str,
    template_name: str,
    language_code: str,
    image_id: str,
    sleep_seconds: float,
):
    url = f"https://graph.facebook.com/{GRAPH_VERSION}/{phone_number_id}/messages"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json",
    }

    results = []
    total = len(contacts_df)

    for idx, row in contacts_df.iterrows():
        raw_number = row.get(phone_column)
        if pd.isna(raw_number):
            continue

        to_number = str(raw_number).strip()
        if not to_number or to_number.lower() == "nan":
            continue

        payload = build_template_payload(
            to_number=to_number,
            template_name=template_name,
            language_code=language_code,
            image_id=image_id,
        )

        try:
            response = requests.post(url, headers=headers, json=payload, timeout=20)
            data = safe_json(response)
            status = "ok" if response.ok else f"err:{response.status_code}"
            print(f"[{idx + 1}/{total}] -> {to_number}: {status}")
            if not response.ok:
                print(f"  Meta error body: {json.dumps(data, ensure_ascii=True)}")
            results.append(
                {
                    "number": to_number,
                    "status_code": response.status_code,
                    "status": status,
                    "response": json.dumps(data, ensure_ascii=True),
                }
            )
        except Exception as exc:  # noqa: BLE001
            print(f"[{idx + 1}/{total}] -> {to_number}: exception {exc}")
            results.append(
                {
                    "number": to_number,
                    "status_code": None,
                    "status": "exception",
                    "response": str(exc),
                }
            )

        time.sleep(sleep_seconds)

    return results


def main():
    args = parse_args()

    excel_path = Path(args.excel_path)
    image_path = Path(args.image_path)
    log_path = Path(args.log_path)

    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")
    if not image_path.exists():
        raise FileNotFoundError(f"Image file not found: {image_path}")

    contacts_df = pd.read_excel(excel_path, sheet_name=args.sheet_name)
    if args.phone_column not in contacts_df.columns:
        raise ValueError(
            f"Column '{args.phone_column}' not found in Excel columns: {list(contacts_df.columns)}"
        )

    precheck_phone_number_id(args.phone_number_id, args.access_token)
    image_id = upload_image(args.phone_number_id, args.access_token, image_path)

    results = send_template_messages(
        contacts_df=contacts_df,
        phone_column=args.phone_column,
        phone_number_id=args.phone_number_id,
        access_token=args.access_token,
        template_name=args.template_name,
        language_code=args.language_code,
        image_id=image_id,
        sleep_seconds=args.sleep_seconds,
    )

    pd.DataFrame(results).to_excel(log_path, index=False)
    print(f"Done. Log saved to {log_path}")


if __name__ == "__main__":
    main()
