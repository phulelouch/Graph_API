import requests
import json
import time
import csv
from datetime import datetime, timedelta
import urllib3

# Disable only if you're comfortable with not verifying SSL certs.
# Develop by Phu Tran (Dvuln) also modify time.sleep between each request to avoid rate-limit 
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

def refresh_graph_tokens(
    refresh_token,
    tenant_id,
    client_id,
    resource,
    proxies,
):
    """
    Refresh the Graph tokens using the /token endpoint via Burp proxy.
    
    For production, you may need to handle more parameters 
    (like client_secret, scope, etc.), and manage 
    user/application context carefully.
    """
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    data = {
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "client_id": client_id,
        "scope": f"{resource}/.default",
    }

    # Disabling SSL verification for interception by Burp
    response = requests.post(token_url, data=data, proxies=proxies, verify=False)
    if response.status_code != 200:
        raise Exception(
            f"Token refresh failed with code {response.status_code}: {response.text}"
        )
    return response.json()


def get_updatable_groups(
    tokens,
    graph_api_endpoint="https://graph.microsoft.com/v1.0/groups",
    estimate_access_endpoint="https://graph.microsoft.com/beta/roleManagement/directory/estimateAccess",
    refresh_token=None,
    tenant_id=None,
    client="MSGraph",
    client_id="d3590ed6-52b3-4102-aeff-aad2292ab01c",
    resource="https://graph.microsoft.com",
    device="Windows",
    browser="Edge",
    output_file="Updatable_groups.csv",
    auto_refresh=False,
    refresh_interval=600,
):
    """
    Python version of the Get-UpdatableGroups functionality, 
    with proxy support for Burp Suite.

    :param tokens: dict, containing at least 'access_token' and optionally 'refresh_token'
    :param graph_api_endpoint: str, initial Microsoft Graph groups endpoint
    :param estimate_access_endpoint: str, endpoint for checking if 'microsoft.directory/groups/members/update' is allowed
    :param refresh_token: str, optional refresh token (if not provided in tokens)
    :param tenant_id: str, tenant ID for token refresh
    :param client: str, client type (MSGraph, etc.) - used here just for reference/logging
    :param client_id: str, client ID used for token refresh
    :param resource: str, resource for which tokens are valid (default: MS Graph)
    :param device: str, device type for logging/metadata
    :param browser: str, browser for logging/metadata
    :param output_file: str, path to CSV output
    :param auto_refresh: bool, if True, refresh token automatically after refresh_interval
    :param refresh_interval: int, time in seconds to wait before refreshing the token

    :return: A list of dictionaries describing updatable groups.
    """

    # Configure the Burp proxy (HTTP & HTTPS):
    proxies = {
        "http": "http://127.0.0.1:8080",
        "https": "http://127.0.0.1:8080",
    }

    # Extract access token
    access_token = tokens.get("access_token")
    if not refresh_token:
        refresh_token = tokens.get("refresh_token")

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    results = []

    print("[*] Now gathering groups and checking if each one is updatable...")

    start_time = datetime.now()
    refresh_delta = timedelta(seconds=refresh_interval)

    next_link = graph_api_endpoint

    while next_link:
        try:
            # Handle token refresh if auto_refresh is enabled and interval has passed
            if auto_refresh and (datetime.now() - start_time) >= refresh_delta:
                print("[*] Pausing script for token refresh...")
                refreshed = refresh_graph_tokens(
                    refresh_token=refresh_token,
                    tenant_id=tenant_id,
                    client_id=client_id,
                    resource=resource,
                    proxies=proxies
                )
                access_token = refreshed["access_token"]
                refresh_token = refreshed.get("refresh_token", refresh_token)
                headers["Authorization"] = f"Bearer {access_token}"

                print("[*] Resuming script...")
                start_time = datetime.now()

            # Get the list of groups via Burp proxy
            response = None
            try:
                response = requests.get(next_link, headers=headers, proxies=proxies, verify=False)
                #print(response.text)
                time.sleep(2)
                # Handle throttling
                if response.status_code == 429:
                    print("[*] Being throttled... sleeping for 5 seconds")
                    time.sleep(5)
                    continue
                response.raise_for_status()
            except requests.exceptions.RequestException as e:
                print(f"[*] Error occurred while retrieving groups: {str(e)}")
                time.sleep(5)
                continue

            response_json = response.json()

            # For each group, check if user can update
            for group in response_json.get("value", []):
                time.sleep(0.5)
                group_id = group.get("id")
                group_name = group.get("displayName", "Unknown")

                # Check if we need to refresh token again (mid-loop):
                if auto_refresh and (datetime.now() - start_time) >= refresh_delta:
                    print("[*] Pausing script for token refresh...")
                    refreshed = refresh_graph_tokens(
                        refresh_token=refresh_token,
                        tenant_id=tenant_id,
                        client_id=client_id,
                        resource=resource,
                        proxies=proxies
                    )
                    access_token = refreshed["access_token"]
                    refresh_token = refreshed.get("refresh_token", refresh_token)
                    headers["Authorization"] = f"Bearer {access_token}"
                    print("[*] Resuming script...")
                    start_time = datetime.now()

                # Prepare request body for estimateAccess
                request_body = {
                    "resourceActionAuthorizationChecks": [
                        {
                            "directoryScopeId": f"/{group_id}",
                            "resourceAction": "microsoft.directory/groups/members/update"
                        }
                    ]
                }

                try:
                    estimate_resp = requests.post(
                        estimate_access_endpoint,
                        headers=headers,
                        data=json.dumps(request_body),
                        proxies=proxies,
                        verify=False
                    )

                    # Handle throttling
                    if estimate_resp.status_code == 429:
                        print("[*] Being throttled... sleeping for 5 seconds")
                        time.sleep(5)
                        continue

                    estimate_resp.raise_for_status()
                    estimate_json = estimate_resp.json()

                    # FIX: "value" is a list, so we index [0]
                    if ("value" in estimate_json
                        and len(estimate_json["value"]) > 0
                        and "accessDecision" in estimate_json["value"][0]
                        and estimate_json["value"][0]["accessDecision"] == "allowed"):
                        
                        print(f"[+] Found updatable group: {group_name} : {group_id}")
                        results.append({
                            "displayName": group.get("displayName"),
                            "id": group.get("id"),
                            "description": group.get("description"),
                            "isAssignableToRole": group.get("isAssignableToRole"),
                            "onPremisesSyncEnabled": group.get("onPremisesSyncEnabled"),
                            "mail": group.get("mail"),
                            "createdDateTime": group.get("createdDateTime"),
                            "visibility": group.get("visibility")
                        })

                except requests.exceptions.RequestException as e:
                    print(f"Error estimating access for group {group_id}: {str(e)}")
                except KeyError as e:
                    print(f"Missing expected key in response for group {group_id}: {str(e)}")

            # Check if there's a next page
            if "@odata.nextLink" in response_json:
                next_link = response_json["@odata.nextLink"]
                print("[*] Processing more groups...")
            else:
                next_link = None

        except Exception as e:
            print(f"Error fetching group IDs: {str(e)}")
            break

    # Summarize
    if results:
        print(f"[*] Found {len(results)} groups that can be updated.")
        for r in results:
            print("=" * 80)
            print(r)
            print("=" * 80)
    else:
        print("[*] No updatable groups found.")

    # Write results to CSV
    if output_file and results:
        fieldnames = [
            "displayName", "id", "description", "isAssignableToRole",
            "onPremisesSyncEnabled", "mail", "createdDateTime", "visibility"
        ]
        try:
            with open(output_file, mode="w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                for res in results:
                    writer.writerow(res)
            print(f"[*] Exported updatable groups to {output_file}")
        except Exception as e:
            print(f"Error writing to CSV file {output_file}: {str(e)}")

    return results


if __name__ == "__main__":
    # Example usage:
    tokens_example = {
        "access_token": "",
        "refresh_token": ""
    }

    updatable = get_updatable_groups(
        tokens=tokens_example,
        tenant_id="",
        auto_refresh=True,
        refresh_interval=600
    )
