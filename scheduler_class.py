import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
from datetime import datetime, timedelta, date
import calendar
from ortools.sat.python import cp_model
import uuid
import imaplib
import email
import time
import os
import smtplib
from email.message import EmailMessage
import config

class Scheduler:
    def __init__(self, group=None, dry_run=False):
        self.group = group
        self.dry_run = dry_run
        self.sheet = self._connect_to_sheet()
        if self.sheet:
            self.employees_df, self.shifts_df, self.requests_df, self.official_schedule_df, self.sandbox_df = self._read_data()

    def _connect_to_sheet(self):
        """Connects to the Google Sheet."""
        try:
            scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
            creds = ServiceAccountCredentials.from_json_keyfile_name('creds.json', scope)
            client = gspread.authorize(creds)
            spreadsheet = client.open(config.SHEET_NAME)
            print("✅ Successfully connected to the Google Sheet.")
            return spreadsheet
        except Exception as e:
            print(f"❌ Error connecting to Google Sheet: {e}")
            return None

    def _read_data(self):
        """Reads all required data from the Google Sheet into DataFrames."""
        print("Reading data from all tabs...")
        employees_ws = self.sheet.worksheet(config.EMPLOYEES_TAB)
        employees_df = pd.DataFrame(employees_ws.get_all_records())
        if self.group:
            employees_df = employees_df[employees_df[config.COL_EMPLOYEE_ROLE] == self.group]

        shifts_ws = self.sheet.worksheet(config.SHIFTS_TAB)
        shifts_df = pd.DataFrame(shifts_ws.get_all_records())

        requests_ws = self.sheet.worksheet(config.REQUESTS_TAB)
        requests_df = pd.DataFrame(requests_ws.get_all_records())
        if not requests_df.empty:
            requests_df = requests_df.rename(columns={
                config.COL_REQUEST_NAME: 'Employee_Name',
                config.COL_REQUEST_START: 'Start_Date',
                config.COL_REQUEST_END: 'End_Date',
                config.COL_REQUEST_TOKENS: 'Tokens_Bid'
            })

        official_schedule_ws = self.sheet.worksheet(config.OFFICIAL_SCHEDULE_TAB)
        official_schedule_df = pd.DataFrame(official_schedule_ws.get_all_records())

        sandbox_schedule_ws = self.sheet.worksheet(config.SANDBOX_SCHEDULE_TAB)
        sandbox_schedule_df = pd.DataFrame(sandbox_schedule_ws.get_all_records())

        return employees_df, shifts_df, requests_df, official_schedule_df, sandbox_schedule_df

    def _read_offers_data(self):
        """Reads only the Offers tab into a DataFrame."""
        print("Reading offers data...")
        offers_ws = self.sheet.worksheet(config.OFFERS_TAB)
        return pd.DataFrame(offers_ws.get_all_records())

    def generate_schedule(self):
        """
        Generates a schedule with ALL features: dynamic dates, multi-day requests,
        locking past days, and using the official schedule as a hint.
        """
        print("--- Starting Schedule Generation (Full-Featured) ---")

        today = datetime.now()
        _, num_days = calendar.monthrange(today.year, today.month)
        today_index = today.day - 1
        print(f"✅ Detected {num_days} days for the current month. Locking all days up to and including Day {today.day}.")

        requests = []
        if not self.requests_df.empty:
            for _, row in self.requests_df.iterrows():
                start_date = pd.to_datetime(row['Start_Date'], dayfirst=True)
                end_date = pd.to_datetime(row['End_Date'], dayfirst=True)
                num_request_days = (end_date - start_date).days + 1
                tokens_per_day = row['Tokens_Bid'] // num_request_days if num_request_days > 0 else 0
                for day_delta in range(num_request_days):
                    current_date = start_date + timedelta(days=day_delta)
                    requests.append((row['Employee_Name'], current_date.day, 'OFF', tokens_per_day))

        all_employees = self.employees_df['Employee_Name'].tolist()
        employees = {'Infirmier': [], 'Intérimaire': []}
        for _, row in self.employees_df.iterrows():
            name, role = row['Employee_Name'], row['Role']
            if role in employees: employees[role].append(name)

        shifts = {}
        for _, row in self.shifts_df.iterrows():
            shift_id = row[config.COL_SHIFT_ID]
            applicable_days = [int(day) for day in str(row[config.COL_SHIFT_DAYS])]
            shifts[shift_id] = {'duration': int(row[config.COL_SHIFT_DURATION] * 100), 'role': row[config.COL_SHIFT_ROLE], 'days': applicable_days}

        days_of_week = [d % 7 for d in range(num_days)]
        model = cp_model.CpModel()
        works = {}
        for e in all_employees:
            for s_id, s_info in shifts.items():
                if s_info['role'] in self.employees_df[self.employees_df['Employee_Name'] == e]['Role'].values:
                    for d in range(num_days):
                        if days_of_week[d] in s_info['days']:
                            works[(e, s_id, d)] = model.NewBoolVar(f'works_{e}_{s_id}_{d}')

        date_columns = self.official_schedule_df.columns[1:]
        for d in range(today_index + 1):
            if d < len(date_columns):
                day_col = date_columns[d]
                for _, row in self.official_schedule_df.iterrows():
                    shift_id, official_employee = row[config.COL_SCHEDULE_SHIFT], row[day_col]
                    if official_employee and (official_employee in all_employees):
                        if (official_employee, shift_id, d) in works:
                            model.Add(works[(official_employee, shift_id, d)] == 1)

        for s_id, s_info in shifts.items():
            for d in range(num_days):
                if days_of_week[d] in s_info['days']:
                    model.AddExactlyOne(works.get((e, s_id, d), 0) for e in employees[s_info['role']])
        for e in all_employees:
            for d in range(num_days): model.AddAtMostOne(works.get((e, s_id, d), 0) for s_id in shifts)
        for e in all_employees:
            if e != 'INT1':
                for d in range(num_days - 6):
                    worked_days = [works[key] for key in works if key[0] == e and d <= key[2] < d + 7]
                    model.Add(sum(worked_days) <= 6)

        request_bonuses = []
        if requests:
          for emp, day, shift_type, penalty in requests:
              day_index = day - 1
              if shift_type == 'OFF':
                  is_working_on_day = [works[key] for key in works if key[0] == emp and key[2] == day_index]
                  request_fulfilled = model.NewBoolVar(f'request_{emp}_{day_index}')
                  model.Add(sum(is_working_on_day) == 0).OnlyEnforceIf(request_fulfilled)
                  model.Add(sum(is_working_on_day) > 0).OnlyEnforceIf(request_fulfilled.Not())
                  request_bonuses.append(penalty * request_fulfilled)

        hint_bonuses = []
        for d in range(today_index + 1, num_days):
            if d < len(date_columns):
                day_col = date_columns[d]
                for _, row in self.official_schedule_df.iterrows():
                    shift_id, official_employee = row[config.COL_SCHEDULE_SHIFT], row[day_col]
                    if official_employee and (official_employee in all_employees):
                        if (official_employee, shift_id, d) in works:
                            hint_bonuses.append(works[(official_employee, shift_id, d)])

        model.Maximize(sum(request_bonuses) + sum(hint_bonuses))

        solver = cp_model.CpSolver()
        solver.parameters.num_search_workers = config.NUM_PARALLEL_WORKERS
        solver.parameters.max_time_in_seconds = config.SOLVER_TIME_LIMIT_SECONDS
        status = solver.Solve(model)

        if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            print("✅ Schedule generated successfully.")
            solution = {}
            for (e, s, d), var in works.items():
                if solver.Value(var):
                    solution[(s, d)] = e
            return solution
        else:
            print("❌ No solution found.")
            return None

    def create_and_send_offers(self, solution):
        print("--- Creating and Sending Schedule Change Offers ---")
        sandbox_data = {}
        date_columns = self.official_schedule_df.columns[1:]
        for col in date_columns:
            sandbox_data[col] = []
            day_index = date_columns.get_loc(col)
            for shift_id in self.official_schedule_df[config.COL_SCHEDULE_SHIFT]:
                employee = solution.get((shift_id, day_index), '')
                sandbox_data[col].append(employee)
        sandbox_df = pd.DataFrame(sandbox_data, index=self.official_schedule_df[config.COL_SCHEDULE_SHIFT])

        winning_request = None
        if not self.requests_df.empty:
            for _, request in self.requests_df.iterrows():
                emp, day = request['Employee_Name'], pd.to_datetime(request['Start_Date'], dayfirst=True).day
                day_col = self.official_schedule_df.columns[day]
                is_working = any(sandbox_df.loc[shift_id, day_col] == emp for shift_id in sandbox_df.index)
                if not is_working:
                    winning_request = request
                    break # Assume first winner is the one

        requester_name = winning_request['Employee_Name'] if winning_request is not None else "SYSTEM"
        token_reward = winning_request['Tokens_Bid'] if winning_request is not None else 0

        sender_email = os.environ.get('GMAIL_ADDRESS')
        app_password = os.environ.get('GMAIL_APP_PASSWORD')
        if not sender_email or not app_password:
            print("❌ Email credentials not found. Cannot send offers.")
            return sandbox_df

        offers_ws = self.sheet.worksheet(config.OFFERS_TAB)
        offers_to_log = []

        all_changes = {}
        for day_col in date_columns:
            for shift_id in self.official_schedule_df[config.COL_SCHEDULE_SHIFT]:
                official_employee = self.official_schedule_df.loc[self.official_schedule_df[config.COL_SCHEDULE_SHIFT] == shift_id, day_col].iloc[0]
                sandbox_employee = sandbox_df.loc[shift_id, day_col]
                official_employee = '' if pd.isna(official_employee) else official_employee
                sandbox_employee = '' if pd.isna(sandbox_employee) else sandbox_employee

                if official_employee != sandbox_employee:
                # Change involving the employee being removed from the shift
                    if official_employee:
                        if official_employee not in all_changes: all_changes[official_employee] = []
                        change_desc = f"On {day_col}, your shift '{shift_id}' was reassigned to {sandbox_employee or 'unassigned'}."
                        all_changes[official_employee].append(change_desc)

                # Change involving the employee being added to the shift
                if sandbox_employee:
                # If the slot was previously occupied, it's a real offer.
                    if official_employee:
                        if sandbox_employee not in all_changes: all_changes[sandbox_employee] = []
                        change_desc = f"On {day_col}, you were assigned to shift '{shift_id}' (previously {official_employee or 'unassigned'})."
                        all_changes[sandbox_employee].append(change_desc)
                    else:
                        # The slot was empty, so this is a "free move". No offer needed.
                        print(f"INFO: Auto-approving free move for {sandbox_employee} to shift '{shift_id}' on {day_col}.")

        for employee_name, changes_list in all_changes.items():
            offer_id = str(uuid.uuid4())
            recipient_email_series = self.employees_df[self.employees_df[config.COL_EMPLOYEE_NAME] == employee_name][config.COL_EMPLOYEE_EMAIL]
            if recipient_email_series.empty:
                print(f"⚠️ Could not find email for {employee_name}. Skipping offer.")
                continue
            recipient_email = recipient_email_series.iloc[0]

            accept_subject = f"ACCEPT-{offer_id}"
            decline_subject = f"DECLINE-{offer_id}"
            accept_link = f"mailto:{config.HR_EMAIL}?subject={accept_subject}"
            decline_link = f"mailto:{config.HR_EMAIL}?subject={decline_subject}"

            email_body = (
                f"Hello {employee_name},\n\n"
            f"To accommodate a request from {requester_name}, you are being offered a reward of {token_reward} tokens to accept the following schedule change. This reward will be given to the first employee who accepts.\n\n"
                + "\n".join(f"- {change}" for change in changes_list)
                + f"\n\nPlease click to accept or decline:\n"
                f"✅ Accept: {accept_link}\n"
                f"❌ Decline: {decline_link}\n\n"
                f"This offer is valid for 1 hour. Offer ID: {offer_id}"
            )

            self._send_email(recipient_email, 'Schedule Change Proposal', email_body)
            expiry_time = (datetime.now() + timedelta(hours=1)).strftime('%Y-%m-%d %H:%M:%S')
            offers_to_log.append([offer_id, employee_name, "PENDING", expiry_time, requester_name])

        if offers_to_log and not self.dry_run:
            offers_ws.append_rows(offers_to_log)
            print(f"✅ Logged {len(offers_to_log)} new offers to the 'Offers' tab.")
        elif offers_to_log:
            print(f"DRY RUN: Would have logged {len(offers_to_log)} new offers.")

        if not self.dry_run:
            sandbox_ws = self.sheet.worksheet(config.SANDBOX_SCHEDULE_TAB)
            sandbox_ws.update([sandbox_df.columns.values.tolist()] + sandbox_df.reset_index().values.tolist())
            print("✅ Sandbox_Schedule tab has been updated.")
        else:
            print("DRY RUN: Would have updated the Sandbox_Schedule tab.")

        return sandbox_df

    def process_email_replies(self):
        print("--- Processing Email Replies ---")
        accepted_count = 0
        declined_count = 0
        try:
            hr_email = os.environ.get('GMAIL_ADDRESS')
            app_password = os.environ.get('GMAIL_APP_PASSWORD')
            if not hr_email or not app_password:
                print("❌ HR Email credentials not found. Cannot process replies.")
                return 0, 0

            offers_ws = self.sheet.worksheet(config.OFFERS_TAB)

            mail = imaplib.IMAP4_SSL("imap.gmail.com")
            mail.login(hr_email, app_password)
            mail.select("inbox")

            status, data_accept = mail.search(None, '(UNSEEN SUBJECT "ACCEPT-")')
            status, data_decline = mail.search(None, '(UNSEEN SUBJECT "DECLINE-")')

            all_ids = data_accept[0].split() + data_decline[0].split()

            for num in all_ids:
                status, data = mail.fetch(num, '(RFC822)')
                msg = email.message_from_bytes(data[0][1])
                subject = msg['subject']

                try:
                    response, offer_id = subject.split('-')
                    if response.upper() == 'ACCEPT':
                        accepted_count += 1
                    elif response.upper() == 'DECLINE':
                        declined_count += 1

                    cell = offers_ws.find(offer_id)
                    if cell:
                        if not self.dry_run:
                            offers_ws.update_cell(cell.row, 4, response.upper())
                            mail.store(num, '+FLAGS', '\\Seen')

                        print(f"✅ Processed reply for Offer {offer_id}. Status set to {response.upper()}.")

                        # Send confirmation email
                        employee_name = offers_ws.cell(cell.row, 2).value
                        employee_email = self.employees_df[self.employees_df[config.COL_EMPLOYEE_NAME] == employee_name][config.COL_EMPLOYEE_EMAIL].iloc[0]
                        confirmation_subject = "Your Response Has Been Recorded"
                        confirmation_body = f"Thank you, your response ('{response.upper()}') for Offer ID {offer_id} has been successfully recorded."
                        self._send_email(employee_email, confirmation_subject, confirmation_body)
                except ValueError:
                    print(f"⚠️ Could not parse subject: '{subject}'. Skipping.")
                    continue

        except Exception as e:
            print(f"❌ An error occurred while processing email replies: {e}")

        return accepted_count, declined_count

    def finalize_schedule(self, sandbox_df):
        print("--- Finalizing Sandbox Schedule Based on Responses ---")
        offers_ws = self.sheet.worksheet(config.OFFERS_TAB)
        offer_data = pd.DataFrame(offers_ws.get_all_records())

        failed_offers = offer_data[offer_data['Status'].isin(['DECLINED', 'PENDING'])]

        if failed_offers.empty:
            print("✅ All offers were accepted. Sandbox is ready for approval.")
            return sandbox_df

        print(f"Found {len(failed_offers)} declined or expired offers. Reverting changes...")

        final_sandbox_df = sandbox_df.copy()

        failed_requesters = set()

        for _, offer in failed_offers.iterrows():
            original_requester = offer[config.COL_OFFER_REQUESTER]
            failed_requesters.add(original_requester)

            for day_col in self.official_schedule_df.columns[1:]:
                 for shift_id in self.official_schedule_df[config.COL_SCHEDULE_SHIFT]:
                    official_employee = self.official_schedule_df.loc[self.official_schedule_df[config.COL_SCHEDULE_SHIFT] == shift_id, day_col].iloc[0]
                    sandbox_employee = sandbox_df.loc[shift_id, day_col]

                    if sandbox_employee != official_employee:
                        final_sandbox_df.loc[shift_id, day_col] = official_employee

        if not self.dry_run:
            sandbox_ws = self.sheet.worksheet(config.SANDBOX_SCHEDULE_TAB)
            sandbox_ws.update([final_sandbox_df.columns.values.tolist()] + final_sandbox_df.reset_index().values.tolist())
            print("✅ Sandbox schedule has been updated with reverted changes.")
        else:
            print("DRY RUN: Would have updated the Sandbox_Schedule tab with reverted changes.")

        for requester in failed_requesters:
            print(f"--> Notifying {requester} that their request could not be fulfilled.")
            # In a real implementation, an email would be sent here.
            # This would also be wrapped in `if not self.dry_run:`.

        return final_sandbox_df

    def redistribute_tokens(self, final_sandbox_df):
        """
        Redistributes tokens from request winners to the first employee who accepts
        an offer to cover the shift(s).
        """
        print("--- Starting Token Redistribution (First-Come, First-Served) ---")

        employees_ws = self.sheet.worksheet(config.EMPLOYEES_TAB)
        emp_data = employees_ws.get_all_records()
        token_balances = {row[config.COL_EMPLOYEE_NAME]: row[config.COL_EMPLOYEE_TOKENS] for row in emp_data}

        offers_ws = self.sheet.worksheet(config.OFFERS_TAB)
        offer_data = pd.DataFrame(offers_ws.get_all_records())

        # Find winning requests based on the final, confirmed schedule
        winners = []
        if not self.requests_df.empty:
            for _, request in self.requests_df.iterrows():
                emp, day = request['Employee_Name'], pd.to_datetime(request['Start_Date'], dayfirst=True).day
                day_col = self.official_schedule_df.columns[day]
                is_working = any(final_sandbox_df.loc[shift_id, day_col] == emp for shift_id in final_sandbox_df.index)
                if not is_working:
                    winners.append(request)

        if not winners:
            print("No winning requests to process for token redistribution.")
            return

        for winner in winners:
            winner_name = winner['Employee_Name']
            tokens_to_distribute = winner['Tokens_Bid']

            # Find offers associated with this winner's request
            related_offers = offer_data[offer_data[config.COL_OFFER_REQUESTER] == winner_name]
            accepted_offers = related_offers[related_offers[config.COL_OFFER_STATUS] == 'ACCEPTED']

            if not accepted_offers.empty:
                # Find the first employee who accepted
                first_responder = accepted_offers.iloc[0][config.COL_OFFER_EMPLOYEE]

                print(f"Processing win for {winner_name}. Awarding {tokens_to_distribute} tokens to first responder: {first_responder}.")

                # Update token balances
                token_balances[winner_name] -= tokens_to_distribute
                token_balances[first_responder] += tokens_to_distribute
            else:
                print(f"No accepted offers found for {winner_name}'s request. No tokens redistributed.")

        # Batch update the token balances in the Google Sheet
        employees_to_update = employees_ws.col_values(1)
        cell_updates = []
        for emp_name, new_balance in token_balances.items():
            try:
                row_index = employees_to_update.index(emp_name) + 1
                cell_updates.append(gspread.Cell(row_index, 4, new_balance))
                cell_updates.append(gspread.Cell(row_index, 6, new_balance))
            except ValueError:
                continue

        if cell_updates and not self.dry_run:
            employees_ws.update_cells(cell_updates)
            print("✅ Token balances have been updated in the Google Sheet.")
        elif cell_updates:
            print("DRY RUN: Would have updated token balances in the Google Sheet.")

    def _send_email(self, recipient_email, subject, body):
        """A helper function to send emails."""
        sender_email = os.environ.get('GMAIL_ADDRESS')
        app_password = os.environ.get('GMAIL_APP_PASSWORD')
        if not sender_email or not app_password:
            print("❌ Email credentials not found. Cannot send email.")
            return False

        if self.dry_run:
            print(f"DRY RUN: Would send email to {recipient_email} with subject '{subject}'.")
            return True

        msg = EmailMessage()
        msg.set_content(body)
        msg['Subject'] = subject
        msg['From'] = sender_email
        msg['To'] = recipient_email

        try:
            server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
            server.login(sender_email, app_password)
            server.send_message(msg)
            server.quit()
            print(f"✅ Email sent successfully to {recipient_email}.")
            return True
        except Exception as e:
            print(f"❌ Failed to send email to {recipient_email}: {e}")
            return False

    def send_hr_summary(self, accepted_count, declined_count):
        print("--- Sending HR Summary Email ---")
        hr_email = config.HR_EMAIL
        subject = "Hourly Schedule Run Summary"
        body = (
            f"The hourly reply processing is complete.\n\n"
            f"- {accepted_count} offers were accepted.\n"
            f"- {declined_count} offers were declined.\n\n"
            "The Sandbox schedule is now finalized and ready for your one-click approval."
        )
        self._send_email(hr_email, subject, body)
