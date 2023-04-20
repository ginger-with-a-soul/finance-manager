from pydrive2.drive import GoogleDrive
import pygsheets as pys
import yaml
from pydrive2.auth import GoogleAuth
from oauth2client.service_account import ServiceAccountCredentials
from pydrive2.auth import AuthError, InvalidConfigError
from pydrive2.drive import GoogleDriveFileList


def login_with_service_account(service: str):
    """
    Google Drive service with a service account.
    note: for the service account to work, you need to share the folder or
    files with the service account email.

    :return: google auth
    """
    # Define the settings dict to use a service account
    # We also can use all options available for the settings dict like
    # oauth_scope,save_credentials,etc.
    settings = {
        "client_config_backend": "service",
        "service_config": {
            "client_json_file_path": service,
        }
    }
    # Create instance of GoogleAuth
    gauth: GoogleAuth = GoogleAuth(settings=settings)

    # Authenticate
    try:
        gauth.ServiceAuth()
    except (AuthError, InvalidConfigError) as e:
        print('Insuccessful google authorization, exiting ...')
        exit(400)

    return gauth


def pysheets_authorize(service_file: str):
    try:
        gc = pys.authorize(service_file=service_file)
    except pys.AuthenticationError as e:
        print('Pysheets auth error ...')
        exit(400)

    return gc


def load_worksheets(google_client, sheet_name: str, sheet_id: str, worksheet_names: list[str]) -> list[pys.Worksheet]:
    worksheets: list[pys.Worksheet] = []

    # Opens a whole google sheet
    spread_sheet: pys.Spreadsheet = google_client.open_by_key(
        sheet_id)

    # Opens individual worksheets that are a part of one google sheet
    # Edit 'parameters.yaml' to change worksheets
    for worksheet_name in worksheet_names:
        worksheet: pys.Worksheet = spread_sheet.worksheet_by_title(
            worksheet_name)
        worksheets.append(worksheet)

    return worksheets


def get_worksheets(service_file: str, month_sheets: dict, worksheet_names: list[str], year_summary_sheet_name: str, year_summary_sheet_id: str, year_summary_worksheet_names: list[str]) -> tuple[dict, dict]:

    google_client = pysheets_authorize(service_file)

    sheet_names: list[str] = list(month_sheets.keys())
    months_worksheets: dict = {}
    year_summary_worksheets: dict = {}

    try:

        # Getting worksheets for months
        for sheet_name in sheet_names:

            months_worksheets[sheet_name] = load_worksheets(
                google_client, sheet_name, month_sheets[sheet_name], worksheet_names)

        year_summary_worksheets[year_summary_sheet_name] = load_worksheets(
            google_client, year_summary_sheet_name, year_summary_sheet_id, year_summary_worksheet_names)

    except pys.SpreadsheetNotFound:
        print('Spread sheet with the given URL missing ...')
        exit(404)

    return months_worksheets, year_summary_worksheets


def load_parameters(filename: str) -> dict:
    with open(filename, "r") as f:
        return yaml.load(f, Loader=yaml.FullLoader)


def find_sheets(folder_id: str, google_auth: GoogleAuth, year_summary_name: str) -> tuple[dict, str]:

    sheets: dict = {}

    drive: GoogleDrive = GoogleDrive(google_auth)
    query: str = f"'{folder_id}' in parents and trashed=false"
    file_list: GoogleDriveFileList = drive.ListFile({'q': query})

    # Id for this sheet is separated from the id's of sheets representing months
    year_summary_sheet_id: str = ''

    for file in file_list.GetList():

        if file['title'] == year_summary_name:
            year_summary_sheet_id = file['id']
        else:
            sheets[file['title']] = file['id']

    return sheets, year_summary_sheet_id


def update_lists(stats: list[tuple[str, float]], address: str, worksheet: pys.Worksheet, reverse: bool = False, num_of_items: int = 0) -> None:

    stats.sort(key=lambda x: x[1], reverse=reverse)
    print(num_of_items, len(stats))

    if num_of_items > len(stats):
        num_of_items = len(stats)

    for i in range(num_of_items):
        # we add 3 because the cells are merged
        value_address: str = address[0] + str(int(address[1:]) + i*3)
        date_address: str = chr(
            ord(address[0]) - 3) + str(int(address[1:]) + i*3)
        worksheet.update_value(value_address, stats[i][1])
        worksheet.update_value(date_address, stats[i][0])


def calculate_statistics(months_worksheets: dict, year_summary_worksheets: dict, params: dict) -> None:

    year_summary_statistics_worksheet: pys.Worksheet = year_summary_worksheets[
        params['year_summary_name']][0]

    money_earned: list[tuple[str, float]] = []
    money_spent: list[tuple[str, float]] = []
    money_saved: list[tuple[str, float]] = []
    most_expensive_items: list[tuple[str, float]] = []
    most_expensive_categories: dict = {}

    total_earned: float = 0
    total_spent: float = 0
    total_saved: float = 0

    for month, worksheets in months_worksheets.items():

        # TODO: handling when values in statistics worksheet are 0 or NA
        # TODO: items and category extraction includes only the highest 1 per month, meaning that
        # for example, if we've spend 10k, 20k on two items and in february we've spent 9k, the two most
        # expensive things will be 20k and 9k, 10k being overshadowed by being 'only' 2nd most expensive
        summary_worksheet: pys.Worksheet = worksheets[0]
        statistics_worksheet: pys.Worksheet = worksheets[2]
        month_name: str = month.split(' ')[1]

        earned: float = float(summary_worksheet.get_value(
            params['actual_income']).split(params['currency_suffix'])[0].replace(',', ''))
        total_earned += earned

        spent: float = float(summary_worksheet.get_value(
            params['actual_expenses']).split(params['currency_suffix'])[0].replace(',', ''))
        total_spent += spent

        saved: float = earned - spent
        total_saved += saved

        most_exp_it_val: float = float(statistics_worksheet.get_value(
            params['most_expensive_item_value']).split(params['currency_suffix'])[0].replace(',', ''))
        most_exp_it_desc: str = statistics_worksheet.get_value(
            params['most_expensive_item_description'])

        most_exp_cat_val: float = float(statistics_worksheet.get_value(
            params['most_expensive_category_value']).split(params['currency_suffix'])[0].replace(',', ''))
        most_exp_cat_desc: str = statistics_worksheet.get_value(
            params['most_expensive_category_description'])

        if most_exp_cat_desc in most_expensive_categories.keys():
            most_expensive_categories[most_exp_cat_desc] += most_exp_cat_val
        else:
            most_expensive_categories[most_exp_cat_desc] = most_exp_cat_val
        most_expensive_items.append(
            (most_exp_it_desc, most_exp_it_val))
        money_earned.append((month_name, earned))
        money_spent.append((month_name, spent))
        money_saved.append((month_name, saved))

    year_summary_statistics_worksheet.update_value(
        params['money_earned'], total_earned)
    year_summary_statistics_worksheet.update_value(
        params['money_spent'], total_spent)
    year_summary_statistics_worksheet.update_value(
        params['money_saved'], total_saved)

    update_lists(
        money_earned, params['top1_most_earned'], year_summary_statistics_worksheet, reverse=True, num_of_items=3)
    update_lists(
        money_spent, params['top1_most_spent'], year_summary_statistics_worksheet, reverse=True, num_of_items=3)
    update_lists(
        money_saved, params['top1_most_saved'], year_summary_statistics_worksheet, reverse=True, num_of_items=3)
    update_lists(
        money_spent, params['top1_least_spent'], year_summary_statistics_worksheet, reverse=False, num_of_items=3)
    update_lists(
        most_expensive_items, params['most_expensive_item'], year_summary_statistics_worksheet, reverse=True, num_of_items=5)
    update_lists(
        list(most_expensive_categories.items()), params['most_expensive_category'], year_summary_statistics_worksheet, reverse=True, num_of_items=5)


def main() -> None:
    parameters: dict = load_parameters("parameters.yaml")

    google_auth: GoogleAuth = login_with_service_account(parameters['service'])

    sheet_names_and_ids, year_summary_sheet_id = find_sheets(
        parameters['folder_id'], google_auth, parameters['year_summary_name'])

    months_worksheets, year_summary_worksheets = get_worksheets(
        parameters["service"], sheet_names_and_ids, parameters['worksheet_names'], parameters['year_summary_name'], year_summary_sheet_id, parameters['year_summary_worksheet_names'])

    calculate_statistics(
        months_worksheets, year_summary_worksheets, parameters)


if __name__ == "__main__":
    main()
