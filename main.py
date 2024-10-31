import csv
import importlib.metadata
import requests
from openpyxl import Workbook
from urllib.parse import urlparse
import os

# GitHub API endpoint for retrieving repository license information
GITHUB_API_URL = "https://api.github.com/repos/{owner}/{repo}/license"

# Function to check if the URL is a GitHub repository


def is_github_repo(url):
    parsed_url = urlparse(url)
    return parsed_url.netloc == "github.com"

# Function to extract owner and repo name from a GitHub URL


def extract_github_repo_info(url):
    parsed_url = urlparse(url)
    parts = parsed_url.path.strip("/").split("/")
    if len(parts) >= 2:
        owner, repo = parts[0], parts[1]
        return owner, repo
    return None, None

# Function to fetch the license from GitHub API


def fetch_license_from_github(homepage_url, github_token=None):
    owner, repo = extract_github_repo_info(homepage_url)
    if not owner or not repo:
        return None

    api_url = GITHUB_API_URL.format(owner=owner, repo=repo)

    headers = {}
    if github_token:
        headers['Authorization'] = f'token {github_token}'

    try:
        response = requests.get(api_url, headers=headers)
        if response.status_code == 200:
            data = response.json()
            license_info = data.get("license", {}).get("spdx_id", "Unknown")
            return license_info
    except requests.RequestException:
        return None

    return None

# Fetch metadata for each installed package


def get_package_metadata(package_name, github_token=None):
    try:
        dist = importlib.metadata.distribution(package_name)
        metadata = dist.metadata
        license_info = metadata.get('License', 'Unknown')  # Fetch license info
        home_page = metadata.get('Home-page', 'Unknown')  # Fetch project URL

        # If the license is 'Unknown', try to fetch it from GitHub
        if license_info.lower() == 'unknown' and home_page.lower() != 'unknown' and is_github_repo(home_page):
            github_license = fetch_license_from_github(home_page, github_token)
            if github_license:
                license_info = github_license

        return {
            'name': package_name,
            'version': dist.version,
            'license': license_info,
            'project_url': home_page
        }
    except importlib.metadata.PackageNotFoundError:
        return None

# Collect metadata for all installed packages


def collect_dependencies(github_token=None):
    dependencies = []
    for dist in importlib.metadata.distributions():
        metadata = get_package_metadata(dist.metadata['Name'], github_token)
        if metadata:
            dependencies.append(metadata)
    return dependencies

# Write dependencies info to an Excel file


def write_to_excel(dependencies, output_file='dependencies.xlsx'):
    wb = Workbook()
    ws = wb.active
    ws.title = "Dependencies"

    # Write header
    headers = ['Package Name', 'Version', 'License', 'Project URL']
    ws.append(headers)

    # Write package information
    for dep in dependencies:
        ws.append([dep['name'], dep['version'], dep['license'], dep['project_url']])

    wb.save(output_file)


# Main function
if __name__ == "__main__":
    # Optional: Use GitHub Personal Access Token for higher rate limits
    github_token = os.getenv("GITHUB_TOKEN")  # You can set it in the environment variables

    deps = collect_dependencies(github_token)
    write_to_excel(deps)
    print(f"Dependency information has been saved to 'dependencies.xlsx'.")
