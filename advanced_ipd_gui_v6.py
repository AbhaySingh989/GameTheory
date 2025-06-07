# Block 1: Imports and Constants
import sys
import os
import json
import random
import math
import datetime
import uuid
import itertools
from collections import defaultdict, Counter

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QComboBox, QSpinBox, QSlider, QTabWidget, QTextBrowser, QDialog,
    QFormLayout, QLineEdit, QDialogButtonBox, QListWidget, QListWidgetItem,
    QCheckBox, QTableWidget, QTableWidgetItem, QStatusBar, QSplitter, QMessageBox,
    QGroupBox, QScrollArea, QSizePolicy # Added QSizePolicy
)
from PyQt6.QtGui import QPainter, QColor, QBrush, QPen, QFont, QIcon, QAction
from PyQt6.QtCore import Qt, QTimer, QSize, QPoint, QRect, pyqtSignal

# Matplotlib integration for PyQt6
import matplotlib
matplotlib.use('QtAgg') # Use QtAgg backend
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

# Pandas for Excel logging
import pandas as pd

# Constants and Configuration
PAYOFFS = {
    ('C', 'C'): (3, 3),  # Both Cooperate (Reward)
    ('C', 'D'): (0, 5),  # P1 Cooperates, P2 Defects (Sucker, Temptation)
    ('D', 'C'): (5, 0),  # P1 Defects, P2 Cooperates (Temptation, Sucker)
    ('D', 'D'): (1, 1)   # Both Defect (Punishment)
}
R, S, T, P = 3, 0, 5, 1 # Standard names

COOPERATE = 'C'
DEFECT = 'D'

STATS_FILE = "ipd_stats.json"
CUSTOM_STRATEGIES_FILE = "custom_strategies.json"
LOG_FILE = "ipd_simulation_log_v6.xlsx"

DEFAULT_ROUNDS = 100
DEFAULT_NOISE = 0.0 # 0%
DEFAULT_FORGIVENESS = 0.0 # 0%
MAX_ROUNDS = 1000
MAX_NOISE_FORGIVENESS = 20 # Percentage

# Block 2: Strategy Logic and Metadata

# Global dictionary to hold all strategies
STRATEGIES = {}
PLAYER_STATS = defaultdict(lambda: {'wins': 0, 'losses': 0, 'draws': 0, 'total_score': 0, 'games_played': 0})
CUSTOM_STRATEGIES = {} # Holds definitions of custom strategies only

# --- Built-in Strategy Functions ---
def always_cooperate(my_history, opponent_history, round_number, forgiveness_prob=0.0):
    return COOPERATE

def always_defect(my_history, opponent_history, round_number, forgiveness_prob=0.0):
    return DEFECT

def tit_for_tat(my_history, opponent_history, round_number, forgiveness_prob=0.0):
    if not opponent_history:
        return COOPERATE
    if opponent_history[-1] == DEFECT and random.random() < forgiveness_prob:
        return COOPERATE
    return opponent_history[-1]

def grudger(my_history, opponent_history, round_number, forgiveness_prob=0.0):
    if DEFECT in opponent_history:
        if opponent_history[-1] == DEFECT and random.random() < forgiveness_prob:
           return COOPERATE
        return DEFECT
    return COOPERATE

def random_strategy(my_history, opponent_history, round_number, forgiveness_prob=0.0):
    return random.choice([COOPERATE, DEFECT])

def tit_for_two_tats(my_history, opponent_history, round_number, forgiveness_prob=0.0):
    if len(opponent_history) < 2:
        return COOPERATE
    if opponent_history[-1] == DEFECT and opponent_history[-2] == DEFECT:
        if random.random() < forgiveness_prob:
             return COOPERATE
        return DEFECT
    return COOPERATE

def suspicious_tit_for_tat(my_history, opponent_history, round_number, forgiveness_prob=0.0):
    if not opponent_history:
        return DEFECT
    if opponent_history[-1] == DEFECT and random.random() < forgiveness_prob:
        return COOPERATE
    return opponent_history[-1]

def generous_tit_for_tat_10(my_history, opponent_history, round_number, forgiveness_prob=0.0):
    if not opponent_history:
        return COOPERATE
    base_move = opponent_history[-1]
    if base_move == DEFECT and random.random() < 0.10:
        return COOPERATE
    if opponent_history[-1] == DEFECT and random.random() < forgiveness_prob:
        return COOPERATE
    return base_move

def pavlov(my_history, opponent_history, round_number, forgiveness_prob=0.0):
    if round_number == 0: return COOPERATE
    last_my_move = my_history[-1]
    last_opp_move = opponent_history[-1]
    last_payoff, _ = PAYOFFS[(last_my_move, last_opp_move)]
    if last_payoff in (R, P):
        if last_payoff == P and random.random() < forgiveness_prob:
             return COOPERATE
        return last_my_move
    else:
        return DEFECT if last_my_move == COOPERATE else COOPERATE

def prober(my_history, opponent_history, round_number, forgiveness_prob=0.0):
    if round_number == 0: return DEFECT
    if round_number == 1: return COOPERATE
    if round_number == 2: return COOPERATE
    # Corrected logic: check if *any* defection occurred during the probe phase (rounds 1 and 2)
    # Opponent's history indices will be 0, 1, 2 at the end of round 3 (when this logic matters first)
    if len(opponent_history) >= 3: # Check opponent's history up to round 3
        if opponent_history[1] == COOPERATE and opponent_history[2] == COOPERATE:
             # Opponent cooperated during probe -> TFT
            if not opponent_history: return COOPERATE # Safety check
            if opponent_history[-1] == DEFECT and random.random() < forgiveness_prob:
                 return COOPERATE
            return opponent_history[-1]
        else:
            # Opponent defected during probe -> Always Defect
            return DEFECT
    else:
         # Not enough history yet (shouldn't be reached after round 3)
         return DEFECT # Default to defect if something unexpected happens


def majority(my_history, opponent_history, round_number, forgiveness_prob=0.0):
    if not opponent_history: return COOPERATE
    coop_count = opponent_history.count(COOPERATE)
    defect_count = len(opponent_history) - coop_count
    if coop_count > defect_count: return COOPERATE
    elif defect_count > coop_count:
        if random.random() < forgiveness_prob: return COOPERATE
        return DEFECT
    else: return COOPERATE


# --- Strategy Metadata ---
BUILT_IN_STRATEGIES_META = {
    "cooperate": {"name": "Always Cooperate", "func": always_cooperate, "id": "cooperate", "is_custom": False,"desc": "Always chooses to cooperate.", "pros_cons": "Pros: Simple, good with cooperators.\nCons: Exploitable.", "analogue": "Altruism."},
    "defect": {"name": "Always Defect", "func": always_defect, "id": "defect", "is_custom": False,"desc": "Always chooses to defect.", "pros_cons": "Pros: Exploits cooperators.\nCons: Poor mutual defection.", "analogue": "Aggression."},
    "tit_for_tat": {"name": "Tit for Tat (TFT)", "func": tit_for_tat, "id": "tit_for_tat", "is_custom": False,"desc": "Starts C, mirrors opponent's last move.", "pros_cons": "Pros: Robust, retaliates, forgives.\nCons: Noise sensitivity.", "analogue": "Reciprocity."},
    "grudger": {"name": "Grudger", "func": grudger, "id": "grudger", "is_custom": False,"desc": "Starts C, defects forever after first D.", "pros_cons": "Pros: Strong deterrent.\nCons: Unforgiving.", "analogue": "Zero tolerance."},
    "random": {"name": "Random", "func": random_strategy, "id": "random", "is_custom": False,"desc": "Chooses C/D randomly (50/50).", "pros_cons": "Pros: Unpredictable.\nCons: Inconsistent.", "analogue": "Capriciousness."},
    "tit_for_two_tats": {"name": "Tit for Two Tats", "func": tit_for_two_tats, "id": "tit_for_two_tats", "is_custom": False,"desc": "Starts C, defects only after two consecutive Ds.", "pros_cons": "Pros: More forgiving than TFT.\nCons: Slower to punish.", "analogue": "Second chances."},
    "suspicious_tft": {"name": "Suspicious TFT", "func": suspicious_tit_for_tat, "id": "suspicious_tft", "is_custom": False,"desc": "Starts D, then mirrors opponent's last move.", "pros_cons": "Pros: Avoids first-move sucker.\nCons: Can initiate conflict.", "analogue": "Initial distrust."},
    "generous_tft_10": {"name": "Generous TFT (10%)", "func": generous_tit_for_tat_10, "id": "generous_tft_10", "is_custom": False,"desc": "Like TFT, but 10% chance to C when should D.", "pros_cons": "Pros: Breaks mutual defection.\nCons: Slightly exploitable.", "analogue": "Occasional forgiveness."},
    "pavlov": {"name": "Pavlov (Win-Stay, Lose-Shift)", "func": pavlov, "id": "pavlov", "is_custom": False,"desc": "Repeats last move if payoff was good (R/T), else switches.", "pros_cons": "Pros: Corrects mistakes, exploits AllC.\nCons: Complex cycles possible.", "analogue": "Reinforcement learning."},
    "prober": {"name": "Prober", "func": prober, "id": "prober", "is_custom": False,"desc": "Starts D, C, C. If opponent always C during probe, plays TFT. Else, plays AllD.", "pros_cons": "Pros: Tests opponent.\nCons: Initial D can be bad.", "analogue": "Testing the waters."},
    "majority": {"name": "Majority", "func": majority, "id": "majority", "is_custom": False,"desc": "Plays opponent's most frequent past move (C on tie).", "pros_cons": "Pros: Adapts to overall behavior.\nCons: Slow reaction, exploitable.", "analogue": "Judging by reputation."}
}

# Populate the global STRATEGIES dictionary initially
STRATEGIES.update(BUILT_IN_STRATEGIES_META)

# --- Helper Function for Custom Strategies ---
def create_custom_strategy_function(rules):
    def custom_strategy_logic(my_history, opponent_history, round_number, forgiveness_prob=0.0):
        move = rules.get('default', COOPERATE)
        rule_triggered_defection = False # Track if a rule specifically caused a Defect decision

        # Rule: Opponent's Last Move (Highest priority if history exists)
        if 'opp_last_move' in rules and opponent_history:
            last_opp_move = opponent_history[-1]
            if last_opp_move == COOPERATE and 'C' in rules['opp_last_move']:
                move = rules['opp_last_move']['C']
                if move == DEFECT: rule_triggered_defection = True
            elif last_opp_move == DEFECT and 'D' in rules['opp_last_move']:
                move = rules['opp_last_move']['D']
                if move == DEFECT: rule_triggered_defection = True

        # Rule: Opponent Cooperation Rate < X% (Only if Opp Last Move didn't apply or wasn't defined)
        elif 'opp_coop_lt' in rules and opponent_history: # Check 'elif'
            coop_count = opponent_history.count(COOPERATE)
            total_moves = len(opponent_history)
            coop_rate = (coop_count / total_moves) * 100 if total_moves > 0 else 100
            if coop_rate < rules['opp_coop_lt']['value']:
                move = rules['opp_coop_lt']['move']
                if move == DEFECT: rule_triggered_defection = True

        # Rule: Round Number > Y (Only if previous rules didn't apply)
        elif 'round_gt' in rules and round_number > rules['round_gt']['value']: # Check 'elif'
             move = rules['round_gt']['move']
             if move == DEFECT: rule_triggered_defection = True


        # Apply global forgiveness only if a rule resulted in Defect
        if move == DEFECT and rule_triggered_defection and random.random() < forgiveness_prob:
            # print(f"Custom Rule Forgiveness Applied: Overriding {DEFECT} with {COOPERATE}") # Debug
            return COOPERATE

        return move # Return the determined move (or default if no rules matched)
    return custom_strategy_logic

# Block 3: Game Logic Class
class Game:
    def __init__(self, strategy1_id, strategy2_id, num_rounds, noise_prob=0.0, forgiveness_prob=0.0):
        self.strategy1_id = strategy1_id
        self.strategy2_id = strategy2_id
        self.num_rounds = num_rounds
        self.noise_prob = noise_prob
        self.forgiveness_prob = forgiveness_prob

        self.history1 = []
        self.history2 = []
        self.score1 = 0
        self.score2 = 0

        # Handle potential manual players in sandbox - assign dummy function if ID is 'manual'
        # Ensure we fetch from the global STRATEGIES dictionary
        self.strategy1_func = STRATEGIES.get(strategy1_id, {}).get('func', lambda *args: COOPERATE) # Default C if missing
        self.strategy2_func = STRATEGIES.get(strategy2_id, {}).get('func', lambda *args: COOPERATE)
        self.strategy1_name = STRATEGIES.get(strategy1_id, {}).get('name', strategy1_id) # Use ID as name if missing
        self.strategy2_name = STRATEGIES.get(strategy2_id, {}).get('name', strategy2_id)

    def play_round(self, round_number):
        # This method is primarily for automated games. Sandbox handles moves differently.
        move1_intended = self.strategy1_func(self.history1, self.history2, round_number, self.forgiveness_prob)
        move2_intended = self.strategy2_func(self.history2, self.history1, round_number, self.forgiveness_prob)

        move1_actual = move1_intended
        if random.random() < self.noise_prob: move1_actual = COOPERATE if move1_intended == DEFECT else DEFECT

        move2_actual = move2_intended
        if random.random() < self.noise_prob: move2_actual = COOPERATE if move2_intended == DEFECT else DEFECT

        self.history1.append(move1_actual)
        self.history2.append(move2_actual)

        payoff1, payoff2 = PAYOFFS[(move1_actual, move2_actual)]
        self.score1 += payoff1
        self.score2 += payoff2

        return move1_actual, move2_actual, payoff1, payoff2

    def run_game(self):
        for r in range(self.num_rounds):
            self.play_round(r)
        return self.score1, self.score2, "".join(self.history1), "".join(self.history2)

    def get_round_data(self, round_num):
         if 0 <= round_num < len(self.history1):
             m1 = self.history1[round_num]
             m2 = self.history2[round_num]
             p1, p2 = PAYOFFS[(m1, m2)]
             return m1, m2, p1, p2
         return None, None, None, None

    def get_current_state(self):
         return (self.score1, self.score2,
                 self.history1[-1] if self.history1 else None,
                 self.history2[-1] if self.history2 else None,
                 len(self.history1))

# Block 4: Persistence Functions (JSON)

def save_stats():
    global PLAYER_STATS
    try:
        stats_to_save = {k: dict(v) for k, v in PLAYER_STATS.items()}
        with open(STATS_FILE, 'w') as f:
            json.dump(stats_to_save, f, indent=4)
    except IOError as e:
        print(f"Error saving player stats: {e}")
        # Consider showing a QMessageBox here in a real app

def load_stats():
    global PLAYER_STATS
    PLAYER_STATS = defaultdict(lambda: {'wins': 0, 'losses': 0, 'draws': 0, 'total_score': 0, 'games_played': 0}) # Reset first
    if not os.path.exists(STATS_FILE):
        return # Just use the empty default dict

    try:
        with open(STATS_FILE, 'r') as f:
            loaded_stats = json.load(f)
            for strat_id, stats in loaded_stats.items():
                # Ensure all keys exist, defaulting to 0 if missing from file
                PLAYER_STATS[strat_id] = {
                    'wins': stats.get('wins', 0),
                    'losses': stats.get('losses', 0),
                    'draws': stats.get('draws', 0),
                    'total_score': stats.get('total_score', 0),
                    'games_played': stats.get('games_played', 0)
                }
    except (IOError, json.JSONDecodeError) as e:
        print(f"Error loading player stats: {e}. Starting with fresh stats.")
        # Ensure PLAYER_STATS is reset if loading fails
        PLAYER_STATS = defaultdict(lambda: {'wins': 0, 'losses': 0, 'draws': 0, 'total_score': 0, 'games_played': 0})


def save_custom_strategies():
    global CUSTOM_STRATEGIES
    try:
        strategies_to_save = {}
        # Iterate over CUSTOM_STRATEGIES which should contain the definitions
        for strat_id, data in CUSTOM_STRATEGIES.items():
             strategies_to_save[strat_id] = {
                 'name': data['name'],
                 'desc': data.get('desc', 'N/A'), # Use .get for safety
                 'pros_cons': data.get('pros_cons', 'N/A'),
                 'analogue': data.get('analogue', 'N/A'),
                 'rules': data['rules'] # Rules are essential
             }
        with open(CUSTOM_STRATEGIES_FILE, 'w') as f:
            json.dump(strategies_to_save, f, indent=4)
    except IOError as e:
        print(f"Error saving custom strategies: {e}")
    except Exception as e:
         print(f"Unexpected error saving custom strategies: {e}")


def load_custom_strategies():
    global STRATEGIES, CUSTOM_STRATEGIES
    CUSTOM_STRATEGIES = {} # Clear runtime custom dict first
    # Ensure built-ins are always present by starting fresh
    STRATEGIES = {}
    STRATEGIES.update(BUILT_IN_STRATEGIES_META)

    if not os.path.exists(CUSTOM_STRATEGIES_FILE):
        return # No custom strategies to load

    try:
        with open(CUSTOM_STRATEGIES_FILE, 'r') as f:
            loaded_custom_defs = json.load(f)

        for strat_id, data in loaded_custom_defs.items():
             if 'rules' not in data or 'name' not in data:
                  print(f"Skipping invalid custom strategy '{strat_id}': missing 'rules' or 'name'.")
                  continue

             try:
                 strategy_func = create_custom_strategy_function(data['rules'])
                 full_data = {
                    'name': data['name'], 'func': strategy_func,
                    'desc': data.get('desc', 'No description.'),
                    'pros_cons': data.get('pros_cons', 'N/A'),
                    'analogue': data.get('analogue', 'N/A'),
                    'is_custom': True,
                    'rules': data['rules'], # Store rules for display/saving again
                    'id': strat_id # Ensure ID is present
                 }
                 CUSTOM_STRATEGIES[strat_id] = full_data # Add definition to CUSTOM dict
                 STRATEGIES[strat_id] = full_data       # Add definition to main STRATEGIES dict
             except Exception as e_create:
                  print(f"Error creating function for custom strategy '{strat_id}': {e_create}")


    except (IOError, json.JSONDecodeError) as e:
        print(f"Error reading or parsing custom strategies file ({CUSTOM_STRATEGIES_FILE}): {e}. Only built-in strategies available.")
        CUSTOM_STRATEGIES = {}
        # Ensure STRATEGIES is reset to only built-ins if load fails
        STRATEGIES = {}
        STRATEGIES.update(BUILT_IN_STRATEGIES_META)
    except Exception as e_load:
         print(f"Unexpected error loading custom strategies: {e_load}")
         CUSTOM_STRATEGIES = {}
         STRATEGIES = {}
         STRATEGIES.update(BUILT_IN_STRATEGIES_META)


# Block 5: Excel Logger Class
class ExcelLogger:
    def __init__(self, filename=LOG_FILE):
        self.filename = filename
        # Define headers directly for easier reference
        self.required_sheets = {
            "Match_Log": ["Timestamp", "Sim_Type", "Tournament_ID", "Tournament_Type", "Match_ID",
                          "P1_ID", "P1_Name", "P2_ID", "P2_Name",
                          "Rounds", "Noise_Perc", "Forgiveness_Perc",
                          "P1_Final_Score", "P2_Final_Score", "Winner_ID", "P1_History", "P2_History"],
            "Strategy_Info": ["Strategy_ID", "Name", "Description", "Version", "Is_Custom", "Timestamp"],
            "Tournament_Summary": ["Tournament_ID", "Tournament_Type", "Start_Timestamp", "End_Timestamp",
                                   "Participants_JSON", "Rounds_Per_Game", "Noise_Perc", "Forgiveness_Perc",
                                   "Winner_ID", "Final_Standings_JSON"]
        }
        self._safe_check_and_create_sheets()


    def _safe_check_and_create_sheets(self):
        """ More robust check/creation of Excel file and sheets """
        try:
            if not os.path.exists(self.filename):
                # Create new file with all sheets
                print(f"Creating new log file: {self.filename}")
                with pd.ExcelWriter(self.filename, engine='openpyxl') as writer:
                    for sheet_name, columns in self.required_sheets.items():
                        pd.DataFrame(columns=columns).to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                # File exists, check sheets and headers
                needs_update = False
                all_data = {}
                try:
                    # Read all existing sheets first to preserve data
                    with pd.ExcelFile(self.filename) as xls:
                        for sheet_name in xls.sheet_names:
                            # Only read sheets we care about to avoid issues with unrelated sheets
                            if sheet_name in self.required_sheets:
                                try:
                                    all_data[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)
                                except Exception as e_read:
                                    print(f"Warning: Could not read sheet '{sheet_name}' from existing log: {e_read}. It might be recreated.")
                                    all_data[sheet_name] = pd.DataFrame(columns=self.required_sheets[sheet_name]) # Create empty with correct headers
                                    needs_update = True

                except Exception as e_open:
                    print(f"Warning: Could not open existing Excel file {self.filename} for validation: {e_open}. Will attempt to overwrite.")
                    # If opening fails, plan to create all sheets fresh
                    for sheet_name, columns in self.required_sheets.items():
                        all_data[sheet_name] = pd.DataFrame(columns=columns)
                    needs_update = True


                # Check if required sheets exist and have correct headers
                for sheet_name, columns in self.required_sheets.items():
                    if sheet_name not in all_data:
                        print(f"Creating missing required sheet: {sheet_name}")
                        all_data[sheet_name] = pd.DataFrame(columns=columns)
                        needs_update = True
                    elif list(all_data[sheet_name].columns) != columns:
                        print(f"Warning: Header mismatch in sheet '{sheet_name}'. Recreating sheet header.")
                        # Preserve data if possible, but ensure header is correct
                        existing_data_values = all_data[sheet_name].values.tolist()
                        all_data[sheet_name] = pd.DataFrame(columns=columns)
                        # Attempt to put data back if column count is same (basic heuristic)
                        if len(existing_data_values) > 0 and len(existing_data_values[0]) == len(columns):
                             try:
                                 all_data[sheet_name] = pd.DataFrame(existing_data_values, columns=columns)
                                 print(f"  - Attempted to preserve data for sheet '{sheet_name}'.")
                             except Exception:
                                 print(f"  - Could not preserve data for sheet '{sheet_name}'. It will be empty.")
                        needs_update = True

                # Write back if any changes were needed
                if needs_update:
                    print("Updating Excel file structure...")
                    with pd.ExcelWriter(self.filename, engine='openpyxl', mode='w') as writer: # Overwrite mode needed
                        for sheet_name, df_sheet in all_data.items():
                            if sheet_name in self.required_sheets: # Only write required sheets
                                df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)

        except Exception as e:
             # Catch-all for unexpected errors during file check/creation
             print(f"FATAL: Error during Excel file setup ({self.filename}): {e}")
             raise IOError(f"Failed to initialize Excel logger: {e}") from e


    def log_match(self, sim_type, p1_id, p2_id, rounds, noise_perc, forgiveness_perc,
                  score1, score2, history1, history2, tournament_id=None, tournament_type=None):
        timestamp = datetime.datetime.now().isoformat()
        match_id = str(uuid.uuid4())[:8]
        winner_id = "Draw"
        if score1 > score2: winner_id = p1_id
        elif score2 > score1: winner_id = p2_id

        log_entry = pd.DataFrame([{
            "Timestamp": timestamp, "Sim_Type": sim_type, "Tournament_ID": tournament_id, "Tournament_Type": tournament_type,
            "Match_ID": match_id, "P1_ID": p1_id, "P1_Name": STRATEGIES.get(p1_id, {}).get('name', p1_id),
            "P2_ID": p2_id, "P2_Name": STRATEGIES.get(p2_id, {}).get('name', p2_id), "Rounds": rounds,
            "Noise_Perc": noise_perc, "Forgiveness_Perc": forgiveness_perc, "P1_Final_Score": score1,
            "P2_Final_Score": score2, "Winner_ID": winner_id, "P1_History": history1, "P2_History": history2
        }])

        try:
            # Append using mode='a' and overlay sheet
            with pd.ExcelWriter(self.filename, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                 sheet = writer.book['Match_Log']
                 startrow = sheet.max_row
                 # If startrow is 0 (only possible if sheet was just created/empty), header needs to be written
                 write_header = (startrow == 0 or (startrow == 1 and sheet.cell(row=1, column=1).value is None)) # Check if essentially empty
                 log_entry.to_excel(writer, sheet_name='Match_Log', index=False, header=write_header, startrow=startrow)
        except Exception as e:
            print(f"Error writing to Excel log (Match_Log): {e}")
            print("Attempting recovery write (may overwrite file)...")
            # Fallback: Read all, append in memory, write all (slower but safer)
            try:
                all_sheets = {}
                with pd.ExcelFile(self.filename) as xls:
                     all_sheets = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names if sheet in self.required_sheets}

                match_log_df = all_sheets.get("Match_Log", pd.DataFrame(columns=self.required_sheets["Match_Log"]))
                # Use concat instead of append for modern pandas
                all_sheets["Match_Log"] = pd.concat([match_log_df, log_entry], ignore_index=True)

                with pd.ExcelWriter(self.filename, engine='openpyxl', mode='w') as writer:
                     for sheet_name, df_sheet in all_sheets.items():
                         df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
                print("Recovery write successful.")
            except Exception as e_recover:
                 print(f"Recovery write FAILED: {e_recover}")


    def log_tournament_summary(self, tournament_id, tournament_type, start_time, participants_ids,
                               rounds_per_game, noise_perc, forgiveness_perc, winner_id, final_standings):
        end_time = datetime.datetime.now().isoformat()
        start_time_str = start_time.isoformat()
        participants_json = json.dumps(participants_ids)
        standings_json = json.dumps(final_standings)

        log_entry = pd.DataFrame([{
            "Tournament_ID": tournament_id, "Tournament_Type": tournament_type, "Start_Timestamp": start_time_str,
            "End_Timestamp": end_time, "Participants_JSON": participants_json, "Rounds_Per_Game": rounds_per_game,
            "Noise_Perc": noise_perc, "Forgiveness_Perc": forgiveness_perc, "Winner_ID": winner_id,
            "Final_Standings_JSON": standings_json
        }])

        try:
             with pd.ExcelWriter(self.filename, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                 sheet = writer.book['Tournament_Summary']
                 startrow = sheet.max_row
                 write_header = (startrow == 0 or (startrow == 1 and sheet.cell(row=1, column=1).value is None))
                 log_entry.to_excel(writer, sheet_name='Tournament_Summary', index=False, header=write_header, startrow=startrow)
        except Exception as e:
            print(f"Error writing to Excel log (Tournament_Summary): {e}")
            # Add recovery write similar to log_match if needed

    def update_strategy_info(self, all_strategies):
        timestamp = datetime.datetime.now().isoformat()
        strategy_data = []
        for strat_id, data in all_strategies.items():
            strategy_data.append({
                "Strategy_ID": strat_id, "Name": data['name'], "Description": data['desc'],
                "Version": data.get('version', '1.0'), "Is_Custom": data.get('is_custom', False), # Safe get
                "Timestamp": timestamp
            })
        if not strategy_data: return
        df_new = pd.DataFrame(strategy_data)

        try:
            all_sheets_data = {}
            try:
                 with pd.ExcelFile(self.filename) as xls:
                    # Only read required sheets
                    all_sheets_data = {sheet_name: pd.read_excel(xls, sheet_name=sheet_name)
                                        for sheet_name in xls.sheet_names if sheet_name in self.required_sheets}
            except FileNotFoundError:
                 print(f"Log file {self.filename} not found for updating strategy info. Creating new file.")
                 # Initialize required sheets if file doesn't exist
                 for sheet_name, columns in self.required_sheets.items():
                      all_sheets_data[sheet_name] = pd.DataFrame(columns=columns)
            except Exception as e_read:
                print(f"Error reading Excel file {self.filename} for strategy update: {e_read}. Will attempt overwrite.")
                # Initialize required sheets if read fails
                for sheet_name, columns in self.required_sheets.items():
                      all_sheets_data[sheet_name] = pd.DataFrame(columns=columns)


            # Get existing strategy info or create empty df if missing
            df_existing = all_sheets_data.get("Strategy_Info", pd.DataFrame(columns=self.required_sheets["Strategy_Info"]))

            # Ensure df_existing has correct columns before concat
            if list(df_existing.columns) != self.required_sheets["Strategy_Info"]:
                 print("Warning: Correcting Strategy_Info header before update.")
                 df_existing = pd.DataFrame(columns=self.required_sheets["Strategy_Info"])


            # Combine, deduplicate, keeping latest timestamp
            df_combined = pd.concat([df_existing, df_new], ignore_index=True)
            df_final = df_combined.drop_duplicates(subset=['Strategy_ID'], keep='last')
            all_sheets_data['Strategy_Info'] = df_final # Update the data in our dictionary

            # Write all required sheets back in 'w' mode (overwrite)
            with pd.ExcelWriter(self.filename, engine='openpyxl', mode='w') as writer:
                for sheet_name, df_sheet in all_sheets_data.items():
                    # Only write sheets we know about and have data for
                    if sheet_name in self.required_sheets:
                        df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)

        except Exception as e:
            print(f"FATAL Error updating Strategy_Info sheet: {e}")


# Block 6: GUI Helper Widgets

class VisualizationWidget(QWidget):
    history_bar_clicked = pyqtSignal(int, int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumHeight(200)
        self.setMinimumWidth(400)
        self.game_state = None
        self.history1 = ""
        self.history2 = ""
        self.player1_name = "Player 1"
        self.player2_name = "Player 2"
        self.max_rounds = DEFAULT_ROUNDS
        self.status_text = "Idle"
        self.bar_height = 20
        self.bar_spacing = 2
        self.bar_y_offset = 100
        self.bar_x_margin = 10
        self.setFocusPolicy(Qt.FocusPolicy.ClickFocus) # Allow click events

    def update_data(self, game_state, history1, history2, p1_name, p2_name, max_rounds, status="Running"):
        self.game_state = game_state
        self.history1 = history1
        self.history2 = history2
        self.player1_name = p1_name
        self.player2_name = p2_name
        self.max_rounds = max(1, max_rounds) # Ensure > 0
        self.status_text = status
        self.update()

    def clear_data(self):
         self.game_state = None; self.history1 = ""; self.history2 = ""
         self.player1_name = "Player 1"; self.player2_name = "Player 2"
         self.max_rounds = DEFAULT_ROUNDS; self.status_text = "Idle"
         self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.fillRect(self.rect(), QColor("#FAFAFA")) # Match QMainWindow background

        if not self.game_state:
            painter.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, "No game running.")
            return

        score1, score2, last_move1, last_move2, current_round = self.game_state
        font = QFont("Arial", 11) # Slightly smaller font
        painter.setFont(font)
        painter.setPen(Qt.GlobalColor.black)

        status_line = f"Status: {self.status_text} | Round: {current_round}/{self.max_rounds}"
        painter.drawText(self.bar_x_margin, 20, status_line)

        p1_score_text = f"{self.player1_name}: {score1}"
        p2_score_text = f"{self.player2_name}: {score2}"
        p1_score_width = painter.fontMetrics().horizontalAdvance(p1_score_text)
        p2_score_width = painter.fontMetrics().horizontalAdvance(p2_score_text)

        painter.drawText(self.bar_x_margin, 45, p1_score_text)
        # Adjust P2 score position dynamically to avoid overlap
        p2_score_x = self.width() - p2_score_width - self.bar_x_margin - 20 # Approx indicator width + margin
        painter.drawText(max(self.width() // 2, p1_score_width + 40, p2_score_x) , 45, p2_score_text)


        indicator_size = 15
        p1_last_move_x = self.bar_x_margin + p1_score_width + 10
        p2_last_move_x = max(self.width() // 2, p1_score_width + 40, p2_score_x) + p2_score_width + 10

        y_indicator = 35
        if last_move1:
            color = QColor("#4CAF50") if last_move1 == COOPERATE else QColor("#FF9800")
            painter.setBrush(QBrush(color)); painter.setPen(QPen(color))
            if last_move1 == COOPERATE: painter.drawEllipse(p1_last_move_x, y_indicator, indicator_size, indicator_size)
            else: painter.drawRect(p1_last_move_x, y_indicator, indicator_size, indicator_size)

        if last_move2:
            color = QColor("#4CAF50") if last_move2 == COOPERATE else QColor("#FF9800")
            painter.setBrush(QBrush(color)); painter.setPen(QPen(color))
            if last_move2 == COOPERATE: painter.drawEllipse(p2_last_move_x, y_indicator, indicator_size, indicator_size)
            else: painter.drawRect(p2_last_move_x, y_indicator, indicator_size, indicator_size)

        # History Bars
        if not self.history1 and not self.history2: return

        available_width = self.width() - 2 * self.bar_x_margin
        bar_width = available_width / self.max_rounds if self.max_rounds > 0 else available_width

        y_p1 = self.bar_y_offset
        painter.setPen(Qt.GlobalColor.black) # Reset pen
        painter.drawText(self.bar_x_margin, y_p1 - 5, f"{self.player1_name} History:")

        # Background for Player 1
        # Adjusted background rect to encompass title and some padding
        p1_history_bg_rect = QRect(self.bar_x_margin, y_p1 - 5 - painter.fontMetrics().height(), int(available_width), self.bar_height + 10 + painter.fontMetrics().height())
        painter.fillRect(p1_history_bg_rect, QColor("#E8F5E9"))

        for i, move in enumerate(self.history1):
            x = self.bar_x_margin + i * bar_width
            # Ensure minimum visible width for bars when many rounds
            draw_width = max(1, bar_width - 1) if bar_width > 1 else bar_width
            rect = QRect(int(x), y_p1, int(draw_width) , self.bar_height)

            if move == COOPERATE:
                painter.setBrush(QBrush(QColor("#4CAF50")))
                painter.setPen(QPen(QColor("#388E3C")))
                side = min(int(draw_width), self.bar_height)
                circle_rect = QRect(int(x + (draw_width - side) / 2), int(y_p1 + (self.bar_height - side) / 2), side, side)
                painter.drawEllipse(circle_rect)
            else: # DEFECT
                painter.setBrush(QBrush(QColor("#FF9800")))
                painter.setPen(QPen(QColor("#F57C00")))
                painter.drawRect(rect)

        y_p2 = y_p1 + self.bar_height + self.bar_spacing + 20 + painter.fontMetrics().height() # Adjusted for p1 bg
        painter.setPen(Qt.GlobalColor.black) # Reset pen
        painter.drawText(self.bar_x_margin, y_p2 - 5, f"{self.player2_name} History:")

        # Background for Player 2
        # Adjusted background rect to encompass title and some padding
        p2_history_bg_rect = QRect(self.bar_x_margin, y_p2 - 5 - painter.fontMetrics().height(), int(available_width), self.bar_height + 10 + painter.fontMetrics().height())
        painter.fillRect(p2_history_bg_rect, QColor("#FFF3E0"))

        for i, move in enumerate(self.history2):
            x = self.bar_x_margin + i * bar_width
            draw_width = max(1, bar_width - 1) if bar_width > 1 else bar_width
            rect = QRect(int(x), y_p2, int(draw_width), self.bar_height)

            if move == COOPERATE:
                painter.setBrush(QBrush(QColor("#4CAF50")))
                painter.setPen(QPen(QColor("#388E3C")))
                side = min(int(draw_width), self.bar_height)
                circle_rect = QRect(int(x + (draw_width - side) / 2), int(y_p2 + (self.bar_height - side) / 2), side, side)
                painter.drawEllipse(circle_rect)
            else: # DEFECT
                painter.setBrush(QBrush(QColor("#FF9800")))
                painter.setPen(QPen(QColor("#F57C00")))
                painter.drawRect(rect)

    def mousePressEvent(self, event):
        click_pos = event.pos()
        available_width = self.width() - 2 * self.bar_x_margin
        bar_width = available_width / self.max_rounds if self.max_rounds > 0 else 0
        if bar_width <= 0: return # Avoid division by zero if width is too small

        y_p1 = self.bar_y_offset
        rect_p1_area = QRect(self.bar_x_margin, y_p1, int(available_width), self.bar_height)
        if rect_p1_area.contains(click_pos):
            # Calculate clicked round index based on position
            round_index = math.floor((click_pos.x() - self.bar_x_margin) / bar_width)
            if 0 <= round_index < len(self.history1):
                self.history_bar_clicked.emit(0, round_index)
            return

        y_p2 = y_p1 + self.bar_height + self.bar_spacing + 20
        rect_p2_area = QRect(self.bar_x_margin, y_p2, int(available_width), self.bar_height)
        if rect_p2_area.contains(click_pos):
            round_index = math.floor((click_pos.x() - self.bar_x_margin) / bar_width)
            if 0 <= round_index < len(self.history2):
                self.history_bar_clicked.emit(1, round_index)
            return

        super().mousePressEvent(event)


class MplCanvas(FigureCanvas):
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        self.fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = self.fig.add_subplot(111)
        super().__init__(self.fig)
        self.setParent(parent)
        FigureCanvas.setSizePolicy(self, QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        FigureCanvas.updateGeometry(self)
        try: # Add robust tight_layout call
             self.fig.tight_layout()
        except Exception as e:
             print(f"Warning: Initial tight_layout failed: {e}")


class CustomStrategyDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Define New Custom Strategy")
        self.setMinimumWidth(450)
        self.layout = QVBoxLayout(self)
        self.form_layout = QFormLayout()

        self.name_edit = QLineEdit(); self.desc_edit = QLineEdit()
        self.name_edit.setPlaceholderText("Unique Name (e.g., Cautious Prober)")
        self.desc_edit.setPlaceholderText("Brief description")
        self.form_layout.addRow("Strategy Name*:", self.name_edit)
        self.form_layout.addRow("Description:", self.desc_edit)

        self.rules_group = QGroupBox("Rules (Applied Top-Down; '*' = required)")
        self.rules_layout = QFormLayout()
        self.opp_last_c_combo = QComboBox(); self.opp_last_c_combo.addItems(["(Not Set)", COOPERATE, DEFECT])
        self.opp_last_d_combo = QComboBox(); self.opp_last_d_combo.addItems(["(Not Set)", COOPERATE, DEFECT])
        opp_last_layout = QHBoxLayout(); opp_last_layout.addWidget(QLabel("If Opp Last C ->")); opp_last_layout.addWidget(self.opp_last_c_combo)
        opp_last_layout.addWidget(QLabel(" | If Opp Last D ->")); opp_last_layout.addWidget(self.opp_last_d_combo)
        self.rules_layout.addRow("1. Opponent's Last Move:", opp_last_layout)

        self.opp_coop_lt_spin = QSpinBox(); self.opp_coop_lt_spin.setRange(0, 100); self.opp_coop_lt_spin.setSuffix("%")
        self.opp_coop_lt_combo = QComboBox(); self.opp_coop_lt_combo.addItems([COOPERATE, DEFECT])
        opp_coop_lt_layout = QHBoxLayout(); opp_coop_lt_layout.addWidget(QLabel("If Opp Coop Rate <")); opp_coop_lt_layout.addWidget(self.opp_coop_lt_spin)
        opp_coop_lt_layout.addWidget(QLabel(" ->")); opp_coop_lt_layout.addWidget(self.opp_coop_lt_combo)
        self.opp_coop_lt_check = QCheckBox("Enable"); opp_coop_lt_outer = QHBoxLayout(); opp_coop_lt_outer.addWidget(self.opp_coop_lt_check); opp_coop_lt_outer.addLayout(opp_coop_lt_layout)
        self.rules_layout.addRow("2. Opp Coop Rate:", opp_coop_lt_outer)

        self.round_gt_spin = QSpinBox(); self.round_gt_spin.setRange(0, MAX_ROUNDS)
        self.round_gt_combo = QComboBox(); self.round_gt_combo.addItems([COOPERATE, DEFECT])
        round_gt_layout = QHBoxLayout(); round_gt_layout.addWidget(QLabel("If Round # >")); round_gt_layout.addWidget(self.round_gt_spin)
        round_gt_layout.addWidget(QLabel(" ->")); round_gt_layout.addWidget(self.round_gt_combo)
        self.round_gt_check = QCheckBox("Enable"); round_gt_outer = QHBoxLayout(); round_gt_outer.addWidget(self.round_gt_check); round_gt_outer.addLayout(round_gt_layout)
        self.rules_layout.addRow("3. Round Number:", round_gt_outer)

        self.default_move_combo = QComboBox(); self.default_move_combo.addItems([COOPERATE, DEFECT])
        self.rules_layout.addRow("4. Default Move*:", self.default_move_combo)

        self.rules_group.setLayout(self.rules_layout); self.layout.addLayout(self.form_layout); self.layout.addWidget(self.rules_group)
        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept); self.button_box.rejected.connect(self.reject); self.layout.addWidget(self.button_box)

    def get_strategy_data(self):
        name = self.name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "Input Error", "Strategy Name cannot be empty.")
            return None
        # Sanitize ID more thoroughly
        sanitized_name = ''.join(c for c in name if c.isalnum() or c in ('_', '-')).strip().lower()
        if not sanitized_name:
            QMessageBox.warning(self, "Input Error", "Strategy Name must contain alphanumeric characters.")
            return None
        strategy_id = "custom_" + sanitized_name

        rules = {}; rules['default'] = self.default_move_combo.currentText() # Default is mandatory
        opp_last_c = self.opp_last_c_combo.currentText(); opp_last_d = self.opp_last_d_combo.currentText()
        if opp_last_c != "(Not Set)" or opp_last_d != "(Not Set)":
            rules['opp_last_move'] = {}
            if opp_last_c != "(Not Set)": rules['opp_last_move']['C'] = opp_last_c
            if opp_last_d != "(Not Set)": rules['opp_last_move']['D'] = opp_last_d
        if self.opp_coop_lt_check.isChecked(): rules['opp_coop_lt'] = {'value': self.opp_coop_lt_spin.value(), 'move': self.opp_coop_lt_combo.currentText()}
        if self.round_gt_check.isChecked(): rules['round_gt'] = {'value': self.round_gt_spin.value(), 'move': self.round_gt_combo.currentText()}

        return {"id": strategy_id, "name": name, "desc": self.desc_edit.text().strip(), "is_custom": True, "rules": rules,
                "pros_cons": "Pros/Cons not specified.", "analogue": "Analogue not specified."}


# Custom QTableWidgetItem for numeric sorting
class NumericTableWidgetItem(QTableWidgetItem):
    def __lt__(self, other):
        try:
            # Try to convert text to float for comparison
            return float(self.text()) < float(other.text())
        except (ValueError, TypeError):
             # Fallback to default string comparison if conversion fails
            return super().__lt__(other)

class LeaderboardDialog(QDialog):
     def __init__(self, standings, tournament_type, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Tournament Leaderboard ({tournament_type})")
        self.setMinimumSize(550, 350) # Slightly larger
        layout = QVBoxLayout(self)
        self.table = QTableWidget()
        headers = ["Rank", "Strategy Name", "Total Score", "Games", "Avg Score"]
        self.table.setColumnCount(len(headers)); self.table.setHorizontalHeaderLabels(headers)
        self.table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers); self.table.setAlternatingRowColors(True); self.table.setSortingEnabled(True)
        self.table.setRowCount(len(standings))

        for row, data in enumerate(standings):
            # Use dictionary access for clarity
            rank = data.get("rank", row + 1)
            name = data.get("name", "N/A")
            score = data.get("score", 0)
            games = data.get("games", 0)
            avg_score = data.get("avg_score", 0.0)

            # Use NumericTableWidgetItem for sortable numeric columns
            rank_item = NumericTableWidgetItem(str(rank))
            name_item = QTableWidgetItem(name) # Name is text
            score_item = NumericTableWidgetItem(str(score))
            games_item = NumericTableWidgetItem(str(games))
            avg_score_item = NumericTableWidgetItem(f"{avg_score:.2f}") # Format display, store number via text

            self.table.setItem(row, 0, rank_item)
            self.table.setItem(row, 1, name_item)
            self.table.setItem(row, 2, score_item)
            self.table.setItem(row, 3, games_item)
            self.table.setItem(row, 4, avg_score_item)


        self.table.resizeColumnsToContents(); self.table.horizontalHeader().setStretchLastSection(True)
        # Default sort by Rank (column 0, Ascending)
        self.table.sortItems(0, Qt.SortOrder.AscendingOrder)

        layout.addWidget(self.table)
        close_button = QPushButton("Close"); close_button.clicked.connect(self.accept)
        layout.addWidget(close_button, alignment=Qt.AlignmentFlag.AlignRight)
        self.setLayout(layout)


# Block 7: Main Application Window (IPDSimulatorV6)

class IPDSimulatorV6(QMainWindow):
    def __init__(self):
        super().__init__() # Call parent constructor first
        self.setWindowTitle("Advanced IPD Simulator V6")
        self.setGeometry(100, 100, 1200, 700)

        # --- Initialize State Variables and Settings EARLY ---
        # Ensure these are defined before any create_..._tab methods that might use them
        self.current_game = None
        self.single_game_timer = QTimer(self)
        # Connect timer timeout later after method is defined

        self.single_game_round_counter = 0
        self.single_game_p1_id = None
        self.single_game_p2_id = None
        self.single_game_max_rounds = DEFAULT_ROUNDS

        self.sandbox_game = None
        self.sandbox_p1_manual = False
        self.sandbox_p2_manual = False
        self.sandbox_p1_move = None
        self.sandbox_p2_move = None

        # *** Initialize settings attributes HERE, BEFORE creating tabs ***
        self.global_noise = DEFAULT_NOISE
        self.global_forgiveness = DEFAULT_FORGIVENESS
        # *** End settings initialization ***

        # Try setting icon - ignore if file not found
        icon_path = "ipd_icon.png"
        if os.path.exists(icon_path): self.setWindowIcon(QIcon(icon_path))
        else: print(f"Icon file not found: {icon_path}")

        # QApplication.setStyle("Fusion") # Overridden by stylesheet

        # --- Global Stylesheet ---
        self.setStyleSheet("""
            QMainWindow {
                background-color: #FAFAFA; /* Light grey background */
            }
            QTabWidget::pane {
                border-top: 2px solid #C2C7CB;
            }
            QTabBar::tab {
                background: #E0E0E0; /* Light grey for inactive tabs */
                border: 1px solid #BDBDBD;
                padding: 8px;
                min-width: 100px; /* Adjust as needed */
            }
            QTabBar::tab:selected {
                background: #FFFFFF; /* White for active tab */
                margin-bottom: -1px; /* Make selected tab blend into pane */
            }
            QPushButton {
                background-color: #4CAF50; /* Green */
                color: white;
                border-radius: 4px;
                padding: 6px 12px;
                border: 1px solid #388E3C;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #397D3B;
            }
            QPushButton:disabled {
                background-color: #BDBDBD;
                color: #757575;
                border-color: #9E9E9E;
            }
            QComboBox, QSpinBox, QLineEdit, QTextBrowser, QListWidget, QTableWidget {
                background-color: #FFFFFF;
                border: 1px solid #BDBDBD;
                padding: 4px;
                border-radius: 3px;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #BDBDBD;
                border-radius: 4px;
                margin-top: 6px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 3px;
                background-color: #FAFAFA; /* Match QMainWindow background */
            }
            QLabel {
                color: #333333; /* Dark grey for text */
            }
            QSlider::groove:horizontal {
                border: 1px solid #bbb;
                background: white;
                height: 10px;
                border-radius: 4px;
            }
            QSlider::handle:horizontal {
                background: #4CAF50; /* Green handle */
                border: 1px solid #388E3C;
                width: 18px;
                margin: -4px 0; /* handle is centered on the groove */
                border-radius: 9px;
            }
            QStatusBar {
                background-color: #E0E0E0;
                color: #333333;
            }
        """)

        # --- Central Widget and Layout ---
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QHBoxLayout(self.central_widget)
        self.splitter = QSplitter(Qt.Orientation.Horizontal)

        # --- Left Panel: Controls ---
        self.control_panel = QWidget()
        self.control_layout = QVBoxLayout(self.control_panel)
        self.control_layout.setContentsMargins(5, 5, 5, 5)
        self.tab_widget = QTabWidget()
        self.tab_widget.setMinimumWidth(380)
        self.control_layout.addWidget(self.tab_widget)

        # --- Create tabs (methods defined below) ---
        # Now it's safe to call create_settings_tab because self.global_noise exists
        self.create_single_game_tab()
        self.create_tournament_tab()
        self.create_strategies_tab()
        self.create_settings_tab() # <<< OK NOW
        self.create_sandbox_tab()
        self.create_teaching_aids_tab()
        self.create_analytics_tab()

        # --- Right Panel: Visualization ---
        self.visualization_panel = QWidget()
        self.visualization_layout = QVBoxLayout(self.visualization_panel)
        self.visualization_layout.setContentsMargins(5, 5, 5, 5)
        self.visualization_widget = VisualizationWidget()
        self.visualization_widget.history_bar_clicked.connect(self.show_history_bar_info)
        self.visualization_layout.addWidget(self.visualization_widget)

        # --- Add panels to splitter ---
        self.splitter.addWidget(self.control_panel)
        self.splitter.addWidget(self.visualization_panel)
        self.splitter.setSizes([420, 780])
        self.main_layout.addWidget(self.splitter)

        # --- Status Bar ---
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready.")

        # --- Connect Signals AFTER methods defined ---
        self.single_game_timer.timeout.connect(self.run_single_game_step)

        # --- Load Persistent Data and Populate UI ---
        # load_stats() and load_custom_strategies() are called externally before __init__
        self.update_all_strategy_selectors() # Populate combos/lists
        self.update_strategy_info_display() # Initial info display

        # Update analytics plot - data should be loaded by now
        self.update_analytics_plot()

        print("IPDSimulatorV6 Initialized.") # Debug print


    # --- Icon Helper ---
    def get_icon(self, filename):
        """ Safely gets a QIcon, returns empty if file not found. """
        if os.path.exists(filename):
            return QIcon(filename)
        else:
            # print(f"Icon file not found: {filename}") # Reduce console noise
            return QIcon() # Return empty icon


    # --- Tab Creation Methods ---
    def create_single_game_tab(self):
        tab = QWidget(); layout = QFormLayout(tab); layout.setSpacing(10)
        self.sg_p1_combo = QComboBox(); self.sg_p2_combo = QComboBox()
        self.sg_p1_combo.setToolTip("Select Player 1 strategy."); self.sg_p2_combo.setToolTip("Select Player 2 strategy.")
        layout.addRow("Player 1:", self.sg_p1_combo); layout.addRow("Player 2:", self.sg_p2_combo)
        self.sg_rounds_spin = QSpinBox(); self.sg_rounds_spin.setRange(1, MAX_ROUNDS); self.sg_rounds_spin.setValue(DEFAULT_ROUNDS)
        self.sg_rounds_spin.setToolTip(f"Rounds per game (1-{MAX_ROUNDS})."); layout.addRow("Rounds:", self.sg_rounds_spin)
        self.sg_speed_slider = QSlider(Qt.Orientation.Horizontal); self.sg_speed_slider.setRange(0, 1000); self.sg_speed_slider.setValue(100) # 0=fast, 1000=slow
        self.sg_speed_slider.setToolTip("Visualization delay (slider left=fast, right=slow)."); spd_layout = QHBoxLayout()
        spd_layout.addWidget(QLabel("Fast")); spd_layout.addWidget(self.sg_speed_slider); spd_layout.addWidget(QLabel("Slow"))
        layout.addRow("Speed:", spd_layout)
        self.sg_run_button = QPushButton(self.get_icon("play.png"), " ▶️ Run Game"); self.sg_run_button.setToolTip("Start simulation.")
        self.sg_stop_button = QPushButton(self.get_icon("stop.png"), " ⏹️ Stop Game"); self.sg_stop_button.setToolTip("Stop simulation.")
        self.sg_run_button.clicked.connect(self.start_single_game); self.sg_stop_button.clicked.connect(self.stop_single_game)
        self.sg_stop_button.setEnabled(False); btn_layout = QHBoxLayout(); btn_layout.addWidget(self.sg_run_button); btn_layout.addWidget(self.sg_stop_button)
        layout.addRow(btn_layout)
        self.tab_widget.addTab(tab, "Single Game")

    def create_tournament_tab(self):
        tab = QWidget(); layout = QVBoxLayout(tab); layout.setSpacing(10)
        part_group = QGroupBox("Select Participants"); part_layout = QVBoxLayout()
        self.tourn_participants_list = QListWidget(); self.tourn_participants_list.setSelectionMode(QListWidget.SelectionMode.NoSelection)
        self.tourn_participants_list.setToolTip("Check strategies to include."); part_layout.addWidget(self.tourn_participants_list)
        sel_btns = QHBoxLayout(); sel_all = QPushButton("Select All"); sel_none = QPushButton("Select None")
        sel_all.clicked.connect(self.select_all_participants); sel_none.clicked.connect(self.select_none_participants)
        sel_btns.addWidget(sel_all); sel_btns.addWidget(sel_none); part_layout.addLayout(sel_btns)
        part_group.setLayout(part_layout); layout.addWidget(part_group)

        set_group = QGroupBox("Tournament Settings"); set_layout = QFormLayout()
        self.tourn_type_combo = QComboBox(); self.tourn_type_combo.addItems(["Round Robin", "Elimination", "Group Stage + Knockout"])
        self.tourn_type_combo.setToolTip("Select format."); self.tourn_type_combo.currentIndexChanged.connect(self.update_tournament_options)
        set_layout.addRow("Type:", self.tourn_type_combo)
        self.tourn_rounds_spin = QSpinBox(); self.tourn_rounds_spin.setRange(1, MAX_ROUNDS); self.tourn_rounds_spin.setValue(50)
        self.tourn_rounds_spin.setToolTip("Rounds per game."); set_layout.addRow("Rounds/Game:", self.tourn_rounds_spin)
        self.tourn_groups_label = QLabel("# Groups:"); self.tourn_groups_spin = QSpinBox(); self.tourn_groups_spin.setRange(1, 16); self.tourn_groups_spin.setValue(2)
        self.tourn_groups_spin.setToolTip("Groups for group stage."); set_layout.addRow(self.tourn_groups_label, self.tourn_groups_spin)
        self.tourn_qualifiers_label = QLabel("Qualifiers/Group:"); self.tourn_qualifiers_spin = QSpinBox(); self.tourn_qualifiers_spin.setRange(1, 8); self.tourn_qualifiers_spin.setValue(2)
        self.tourn_qualifiers_spin.setToolTip("Qualifiers per group."); set_layout.addRow(self.tourn_qualifiers_label, self.tourn_qualifiers_spin)
        self.tourn_seeding_check = QCheckBox("Enable Seeding (Alphabetical)"); self.tourn_seeding_check.setToolTip("Sort participants alphabetically for Elimination/Knockout.")
        set_layout.addRow(self.tourn_seeding_check)
        set_group.setLayout(set_layout); layout.addWidget(set_group)

        self.tourn_run_button = QPushButton(self.get_icon("play.png"), " 🏆 Run Tournament"); self.tourn_run_button.setToolTip("Start tournament.")
        self.tourn_run_button.clicked.connect(self.run_tournament); layout.addWidget(self.tourn_run_button, alignment=Qt.AlignmentFlag.AlignCenter)
        self.update_tournament_options(); self.tab_widget.addTab(tab, "Tournament")


    def create_strategies_tab(self):
        tab = QWidget(); layout = QVBoxLayout(tab); layout.setSpacing(10)
        sel_layout = QHBoxLayout(); sel_layout.addWidget(QLabel("Select Strategy:")); self.strat_info_combo = QComboBox()
        self.strat_info_combo.setToolTip("Select strategy to view details."); self.strat_info_combo.currentIndexChanged.connect(self.update_strategy_info_display)
        sel_layout.addWidget(self.strat_info_combo, 1); layout.addLayout(sel_layout) # Add stretch factor
        self.strat_info_browser = QTextBrowser(); self.strat_info_browser.setOpenExternalLinks(True); self.strat_info_browser.setPlaceholderText("Select strategy to view details.")
        layout.addWidget(self.strat_info_browser)
        cust_group = QGroupBox("Custom Strategies"); cust_layout = QVBoxLayout()
        self.define_custom_button = QPushButton(self.get_icon("add.png"), " ✨ Define New"); self.define_custom_button.setToolTip("Create custom strategy.")
        self.define_custom_button.clicked.connect(self.open_define_custom_strategy_dialog); cust_layout.addWidget(self.define_custom_button)
        ls_btns = QHBoxLayout(); self.load_custom_button = QPushButton(self.get_icon("load.png"), " 📂 Load"); self.save_custom_button = QPushButton(self.get_icon("save.png"), "💾 Save")
        self.load_custom_button.setToolTip(f"Load from {CUSTOM_STRATEGIES_FILE}."); self.save_custom_button.setToolTip(f"Save to {CUSTOM_STRATEGIES_FILE}.")
        self.load_custom_button.clicked.connect(self.load_and_update_custom_strategies); self.save_custom_button.clicked.connect(save_custom_strategies)
        ls_btns.addWidget(self.load_custom_button); ls_btns.addWidget(self.save_custom_button); cust_layout.addLayout(ls_btns)
        cust_group.setLayout(cust_layout); layout.addWidget(cust_group)
        self.tab_widget.addTab(tab, "Strategies")

    def create_settings_tab(self):
        tab = QWidget(); layout = QFormLayout(tab); layout.setSpacing(15)
        # Scale is 0-2000 for 0.0 to 20.0% (0.01% precision)
        slider_max = MAX_NOISE_FORGIVENESS * 100

        self.noise_slider = QSlider(Qt.Orientation.Horizontal)
        self.noise_slider.setRange(0, slider_max)
        # Set initial value based on the ALREADY INITIALIZED self.global_noise
        self.noise_slider.setValue(int(self.global_noise * 10000)) # Convert 0.0-0.2 to 0-2000
        self.noise_slider.setToolTip(f"Move flip probability (0-{MAX_NOISE_FORGIVENESS}%).")
        self.noise_label = QLabel(f"{self.global_noise*100:.1f}%") # Display initial value
        self.noise_slider.valueChanged.connect(self.update_noise_setting)
        n_layout = QHBoxLayout(); n_layout.addWidget(self.noise_slider); n_layout.addWidget(self.noise_label)
        layout.addRow("Global Noise:", n_layout)

        self.forgiveness_slider = QSlider(Qt.Orientation.Horizontal)
        self.forgiveness_slider.setRange(0, slider_max)
        # Set initial value based on the ALREADY INITIALIZED self.global_forgiveness
        self.forgiveness_slider.setValue(int(self.global_forgiveness * 10000)) # Convert 0.0-0.2 to 0-2000
        self.forgiveness_slider.setToolTip(f"Chance to forgive defection (0-{MAX_NOISE_FORGIVENESS}%).")
        self.forgiveness_label = QLabel(f"{self.global_forgiveness*100:.1f}%") # Display initial value
        self.forgiveness_slider.valueChanged.connect(self.update_forgiveness_setting)
        f_layout = QHBoxLayout(); f_layout.addWidget(self.forgiveness_slider); f_layout.addWidget(self.forgiveness_label)
        layout.addRow("Global Forgiveness:", f_layout)

        self.tab_widget.addTab(tab, "Settings")


    def create_sandbox_tab(self):
        tab = QWidget(); layout = QVBoxLayout(tab); layout.setSpacing(10)
        p_sel = QFormLayout(); self.sandbox_p1_combo = QComboBox(); self.sandbox_p2_combo = QComboBox()
        # Add items later in update_all_strategy_selectors
        self.sandbox_p1_combo.setToolTip("Select P1."); self.sandbox_p2_combo.setToolTip("Select P2.")
        self.sandbox_p1_combo.currentIndexChanged.connect(self.update_sandbox_ui); self.sandbox_p2_combo.currentIndexChanged.connect(self.update_sandbox_ui)
        p_sel.addRow("Player 1:", self.sandbox_p1_combo); p_sel.addRow("Player 2:", self.sandbox_p2_combo); layout.addLayout(p_sel)

        self.sandbox_p1_manual_group = QGroupBox("P1 Manual Move"); p1_man = QHBoxLayout(); self.sandbox_p1_coop_button = QPushButton("🤝 Cooperate"); self.sandbox_p1_defect_button = QPushButton("😠 Defect")
        self.sandbox_p1_coop_button.clicked.connect(lambda: self.sandbox_manual_move(0, COOPERATE)); self.sandbox_p1_defect_button.clicked.connect(lambda: self.sandbox_manual_move(0, DEFECT))
        p1_man.addWidget(self.sandbox_p1_coop_button); p1_man.addWidget(self.sandbox_p1_defect_button); self.sandbox_p1_manual_group.setLayout(p1_man); self.sandbox_p1_manual_group.setVisible(False); layout.addWidget(self.sandbox_p1_manual_group)

        self.sandbox_p2_manual_group = QGroupBox("P2 Manual Move"); p2_man = QHBoxLayout(); self.sandbox_p2_coop_button = QPushButton("🤝 Cooperate"); self.sandbox_p2_defect_button = QPushButton("😠 Defect")
        self.sandbox_p2_coop_button.clicked.connect(lambda: self.sandbox_manual_move(1, COOPERATE)); self.sandbox_p2_defect_button.clicked.connect(lambda: self.sandbox_manual_move(1, DEFECT))
        p2_man.addWidget(self.sandbox_p2_coop_button); p2_man.addWidget(self.sandbox_p2_defect_button); self.sandbox_p2_manual_group.setLayout(p2_man); self.sandbox_p2_manual_group.setVisible(False); layout.addWidget(self.sandbox_p2_manual_group)

        ctrls = QHBoxLayout(); self.sandbox_next_round_button = QPushButton(self.get_icon("next.png"), " ⏭️ Next Round"); self.sandbox_reset_button = QPushButton(self.get_icon("reset.png"), " 🔄 Reset Sandbox")
        self.sandbox_next_round_button.setToolTip("Advance one round."); self.sandbox_reset_button.setToolTip("Reset sandbox state.")
        self.sandbox_next_round_button.clicked.connect(self.sandbox_play_next_round); self.sandbox_reset_button.clicked.connect(self.sandbox_reset)
        self.sandbox_next_round_button.setEnabled(False); ctrls.addWidget(self.sandbox_next_round_button); ctrls.addWidget(self.sandbox_reset_button)
        layout.addLayout(ctrls); layout.addStretch(); self.tab_widget.addTab(tab, "Sandbox")

    def create_teaching_aids_tab(self):
        tab = QWidget(); layout = QVBoxLayout(tab); layout.setSpacing(10)
        exp_group = QGroupBox("IPD Explanation"); exp_layout = QVBoxLayout()
        help_browser = QTextBrowser(); help_browser.setOpenExternalLinks(True)
        help_browser.setHtml(f"""
            <p>The <strong>Iterated Prisoner's Dilemma (IPD)</strong> is a game theory scenario.</p>
            <p>Two players repeatedly choose: <strong>Cooperate (C)</strong> or <strong>Defect (D)</strong>.</p>
            <p>Payoffs (You, Opponent):</p>
            <ul>
                <li>C vs C: ({R}, {R}) - Reward</li>
                <li>C vs D: ({S}, {T}) - Sucker / Temptation</li>
                <li>D vs C: ({T}, {S}) - Temptation / Sucker</li>
                <li>D vs D: ({P}, {P}) - Punishment</li>
            </ul>
            <p>The 'dilemma': Defecting is individually rational short-term (T>{R}, P>{S}), but mutual defection (P,P) is worse than mutual cooperation (R,R).</p>
            <p>'Iteration' allows strategies based on past moves.</p>
            <p><a href="https://en.wikipedia.org/wiki/Prisoner%27s_dilemma">Wikipedia: Prisoner's Dilemma</a></p>
        """)
        exp_layout.addWidget(help_browser); exp_group.setLayout(exp_layout); layout.addWidget(exp_group)

        scn_group = QGroupBox("Load Scenario Participants"); scn_layout = QHBoxLayout()
        self.scenario_combo = QComboBox(); self.scenario_combo.addItems(["Select Scenario...", "Axelrod's Participants", "All Cooperate vs All Defect", "TFT vs Suspicious TFT", "All Built-in"])
        self.scenario_combo.setToolTip("Pre-select participants on Tournament tab."); scn_layout.addWidget(self.scenario_combo)
        load_btn = QPushButton("Load Scenario"); load_btn.setToolTip("Select participants."); load_btn.clicked.connect(self.load_scenario)
        scn_layout.addWidget(load_btn); scn_group.setLayout(scn_layout); layout.addWidget(scn_group); layout.addStretch()
        self.tab_widget.addTab(tab, "Help/Scenarios")

    def create_analytics_tab(self):
        tab = QWidget(); layout = QVBoxLayout(tab); layout.setSpacing(10)
        ctrls = QHBoxLayout(); ctrls.addWidget(QLabel("Plot Type:")); self.plot_type_combo = QComboBox()
        self.plot_type_combo.addItems(["Total Score", "Win Rate (%)", "Average Score per Game", "H2H Matrix (Placeholder)", "Winners History (Placeholder)"])
        self.plot_type_combo.setToolTip("Select statistic to plot."); ctrls.addWidget(self.plot_type_combo, 1) # Stretch combobox
        self.refresh_plot_button = QPushButton(self.get_icon("refresh.png"), " 📊 Refresh Plot"); self.refresh_plot_button.setToolTip("Update plot.")
        self.refresh_plot_button.clicked.connect(self.update_analytics_plot); ctrls.addWidget(self.refresh_plot_button)
        layout.addLayout(ctrls)
        self.analytics_canvas = MplCanvas(self, width=5, height=4, dpi=100); layout.addWidget(self.analytics_canvas)
        self.tab_widget.addTab(tab, "Analytics")


    # --- Helper and Update Methods Implementation ---
    def update_all_strategy_selectors(self):
        # Get current selections to try and restore them later
        sg_p1_current = self.sg_p1_combo.currentData()
        sg_p2_current = self.sg_p2_combo.currentData()
        strat_info_current = self.strat_info_combo.currentData()
        sandbox_p1_current = self.sandbox_p1_combo.currentData()
        sandbox_p2_current = self.sandbox_p2_combo.currentData()
        tourn_checked = self.get_selected_participants() # Get list of checked IDs


        # Sort strategies by name for display
        sorted_strategies = sorted(STRATEGIES.items(), key=lambda item: item[1]['name'])

        # --- Clear existing items ---
        self.sg_p1_combo.clear()
        self.sg_p2_combo.clear()
        self.strat_info_combo.clear()
        self.tourn_participants_list.clear()
        self.sandbox_p1_combo.clear()
        self.sandbox_p2_combo.clear()

        # --- Repopulate ---
        self.strat_info_combo.addItem("Select Strategy...", userData=None)
        self.sandbox_p1_combo.addItem("Manual Input", userData="manual")
        self.sandbox_p2_combo.addItem("Manual Input", userData="manual")

        for strategy_id, data in sorted_strategies:
            name = data['name']
            self.sg_p1_combo.addItem(name, userData=strategy_id)
            self.sg_p2_combo.addItem(name, userData=strategy_id)
            self.strat_info_combo.addItem(name, userData=strategy_id)
            item = QListWidgetItem(name); item.setData(Qt.ItemDataRole.UserRole, strategy_id)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            # Restore checked state if ID was previously checked
            item.setCheckState(Qt.CheckState.Checked if strategy_id in tourn_checked else Qt.CheckState.Unchecked)
            self.tourn_participants_list.addItem(item)
            self.sandbox_p1_combo.addItem(name, userData=strategy_id)
            self.sandbox_p2_combo.addItem(name, userData=strategy_id)

        # --- Restore selections ---
        def restore_combo(combo, data_to_restore, default_index=0):
            index = combo.findData(data_to_restore)
            combo.setCurrentIndex(index if index >= 0 else default_index)

        restore_combo(self.sg_p1_combo, sg_p1_current, 0)
        # Default second player to different strategy if possible
        default_sg_p2 = 1 if self.sg_p2_combo.count() > 1 and sg_p1_current != self.sg_p2_combo.itemData(1) else 0
        restore_combo(self.sg_p2_combo, sg_p2_current, default_sg_p2)

        restore_combo(self.strat_info_combo, strat_info_current, 0) # Default to "Select..."
        restore_combo(self.sandbox_p1_combo, sandbox_p1_current, 0) # Default to Manual
        restore_combo(self.sandbox_p2_combo, sandbox_p2_current, 0)


    def update_strategy_info_display(self):
        selected_id = self.strat_info_combo.currentData()
        if selected_id and selected_id in STRATEGIES:
            data = STRATEGIES[selected_id]
            # Basic HTML formatting
            info_html = f"<h3>{data.get('name', 'Unknown Name')}</h3>"                         f"<p><b>ID:</b> {selected_id}</p>"                         f"<p><b>Desc:</b> {data.get('desc', 'N/A')}</p>"                         f"<p><b>Pros/Cons:</b><br>{data.get('pros_cons', 'N/A').replace(chr(10), '<br>')}</p>"                         f"<p><b>Analogue:</b><br>{data.get('analogue', 'N/A').replace(chr(10), '<br>')}</p>"                         f"<p><b>Type:</b> {'Custom' if data.get('is_custom', False) else 'Built-in'}</p>"
            if data.get('is_custom', False):
                 # Use .get for rules too, in case of malformed data
                 rules_str = json.dumps(data.get('rules', {}), indent=2)
                 info_html += f"<p><b>Rules:</b><pre>{rules_str}</pre></p>"
            self.strat_info_browser.setHtml(info_html)
        else:
            self.strat_info_browser.clear()
            self.strat_info_browser.setPlaceholderText("Select strategy from dropdown...")

    def update_noise_setting(self, value):
        # Scale is 0-2000 for 0.0 to 20.0%
        self.global_noise = value / 10000.0 # Convert 0-2000 to 0.0-0.2
        self.noise_label.setText(f"{self.global_noise*100:.1f}%")

    def update_forgiveness_setting(self, value):
        # Scale is 0-2000 for 0.0 to 20.0%
        self.global_forgiveness = value / 10000.0 # Convert 0-2000 to 0.0-0.2
        self.forgiveness_label.setText(f"{self.global_forgiveness*100:.1f}%")

    def update_player_stats(self, p1_id, p2_id, score1, score2):
        global PLAYER_STATS
        for pid, score, opp_score in [(p1_id, score1, score2), (p2_id, score2, score1)]:
             # Skip dummy IDs used for manual sandbox players
             if pid.startswith("manual_"): continue
             if pid not in STRATEGIES: # Skip if strategy was removed but stats remain
                 print(f"Warning: Attempting to update stats for unknown strategy ID: {pid}")
                 continue

             # Ensure stats entry exists
             if pid not in PLAYER_STATS:
                 PLAYER_STATS[pid] = {'wins': 0, 'losses': 0, 'draws': 0, 'total_score': 0, 'games_played': 0}

             PLAYER_STATS[pid]['games_played'] += 1
             PLAYER_STATS[pid]['total_score'] += score
             if score > opp_score: PLAYER_STATS[pid]['wins'] += 1
             elif opp_score > score: PLAYER_STATS[pid]['losses'] += 1
             else: PLAYER_STATS[pid]['draws'] += 1

    def show_history_bar_info(self, player_index, round_index):
        game = self.current_game if self.current_game else self.sandbox_game
        if game:
             try:
                m1, m2, p1_payoff, p2_payoff = game.get_round_data(round_index)
                if m1 is not None:
                    p1_name = game.strategy1_name
                    p2_name = game.strategy2_name
                    player_name = p1_name if player_index == 0 else p2_name
                    my_move = m1 if player_index == 0 else m2
                    my_payoff = p1_payoff if player_index == 0 else p2_payoff
                    opp_move = m2 if player_index == 0 else m1
                    opp_name = p2_name if player_index == 0 else p1_name
                    # Include opponent name for clarity
                    msg = f"R{round_index + 1}: {player_name}={my_move}({my_payoff}) vs {opp_name}={opp_move}. Payoffs: ({p1_payoff}, {p2_payoff})"
                    self.status_bar.showMessage(msg, 6000) # Slightly longer display time
             except IndexError:
                 self.status_bar.showMessage(f"Invalid round index: {round_index + 1}", 3000)
             except Exception as e:
                 print(f"Error showing history bar info: {e}")


    # --- Single Game Logic Implementation ---
    def start_single_game(self):
        # Stop previous game first if running
        if self.single_game_timer.isActive():
             self.stop_single_game()

        self.single_game_p1_id = self.sg_p1_combo.currentData()
        self.single_game_p2_id = self.sg_p2_combo.currentData()
        self.single_game_max_rounds = self.sg_rounds_spin.value()
        if not self.single_game_p1_id or not self.single_game_p2_id:
            QMessageBox.warning(self, "Setup Error", "Select strategies."); return

        try:
            # Ensure strategy IDs are valid before creating Game
            if self.single_game_p1_id not in STRATEGIES or self.single_game_p2_id not in STRATEGIES:
                 QMessageBox.critical(self, "Error", "Invalid strategy selected.")
                 return

            self.current_game = Game(self.single_game_p1_id, self.single_game_p2_id, self.single_game_max_rounds, self.global_noise, self.global_forgiveness)
            self.single_game_round_counter = 0
            self.sg_run_button.setEnabled(False)
            self.sg_stop_button.setEnabled(True)
            self.visualization_widget.clear_data() # Clear previous viz
            # Update viz with initial state before starting timer
            self.visualization_widget.update_data(
                self.current_game.get_current_state(), "", "",
                self.current_game.strategy1_name, self.current_game.strategy2_name,
                self.single_game_max_rounds, status="Starting"
            )
            self.status_bar.showMessage(f"Starting: {self.current_game.strategy1_name} vs {self.current_game.strategy2_name}")
            QApplication.processEvents() # Ensure UI updates before timer starts

            # Invert slider: 0=fast (small delay), 1000=slow (1000ms delay)
            delay = self.sg_speed_slider.maximum() - self.sg_speed_slider.value()
            delay = max(5, delay) # Ensure minimum delay > 0
            self.single_game_timer.start(delay)
        except KeyError as e:
             QMessageBox.critical(self, "Error", f"Failed to start game. Unknown strategy ID: {e}")
             self.sg_run_button.setEnabled(True); self.sg_stop_button.setEnabled(False)
        except Exception as e:
             QMessageBox.critical(self, "Error", f"An unexpected error occurred starting the game: {e}")
             self.sg_run_button.setEnabled(True); self.sg_stop_button.setEnabled(False)


    def run_single_game_step(self):
        if not self.current_game:
            self.single_game_timer.stop()
            print("Warning: run_single_game_step called with no current_game.")
            return # Stop if game ended unexpectedly

        if self.single_game_round_counter < self.single_game_max_rounds:
            try:
                self.current_game.play_round(self.single_game_round_counter)
                current_state = self.current_game.get_current_state()
                self.visualization_widget.update_data(current_state, "".join(self.current_game.history1), "".join(self.current_game.history2),
                                                    self.current_game.strategy1_name, self.current_game.strategy2_name,
                                                    self.single_game_max_rounds, status="Running")
                self.single_game_round_counter += 1
                QApplication.processEvents() # Essential to keep UI responsive
            except Exception as e:
                 print(f"Error during single game step {self.single_game_round_counter}: {e}")
                 self.status_bar.showMessage(f"Error in round {self.single_game_round_counter}. Stopping.", 5000)
                 self.stop_single_game() # Stop game on error
        else:
            # Game finished normally after max rounds
            self.stop_single_game()

    def stop_single_game(self):
        self.single_game_timer.stop()
        if self.current_game:
            try:
                s1, s2 = self.current_game.score1, self.current_game.score2
                h1, h2 = "".join(self.current_game.history1), "".join(self.current_game.history2)
                p1_id, p2_id = self.current_game.strategy1_id, self.current_game.strategy2_id
                rounds_played = len(self.current_game.history1) # Actual rounds played

                state = self.current_game.get_current_state()
                self.visualization_widget.update_data(state, h1, h2, self.current_game.strategy1_name, self.current_game.strategy2_name, self.single_game_max_rounds, status="Finished")

                # Update stats only if game completed at least one round
                if rounds_played > 0:
                     self.update_player_stats(p1_id, p2_id, s1, s2)
                     # Log only if game completed normally or stopped after rounds
                     excel_logger.log_match(sim_type="Single Game", p1_id=p1_id, p2_id=p2_id, rounds=rounds_played,
                                           noise_perc=self.global_noise*100, forgiveness_perc=self.global_forgiveness*100, score1=s1, score2=s2, history1=h1, history2=h2)

                winner = "Draw"
                if s1 > s2: winner = self.current_game.strategy1_name
                elif s2 > s1: winner = self.current_game.strategy2_name
                self.status_bar.showMessage(f"Game finished. Score: {s1}-{s2}. Winner: {winner}. Logged.", 5000)

            except Exception as e:
                 print(f"Error during game finalization/logging: {e}")
                 self.status_bar.showMessage("Error finalizing game.", 5000)
            finally:
                 self.current_game = None # Clear finished/stopped game
        else:
             # If stop is called when no game is running (e.g., double click)
             self.status_bar.showMessage("No active game to stop.", 3000)


        # Always reset UI state after stop attempt
        self.sg_run_button.setEnabled(True)
        self.sg_stop_button.setEnabled(False)


    # --- Tournament Logic Implementation ---
    def get_selected_participants(self):
        return [self.tourn_participants_list.item(i).data(Qt.ItemDataRole.UserRole)
                for i in range(self.tourn_participants_list.count())
                if self.tourn_participants_list.item(i).checkState() == Qt.CheckState.Checked]

    def select_all_participants(self):
        for i in range(self.tourn_participants_list.count()): self.tourn_participants_list.item(i).setCheckState(Qt.CheckState.Checked)

    def select_none_participants(self):
        for i in range(self.tourn_participants_list.count()): self.tourn_participants_list.item(i).setCheckState(Qt.CheckState.Unchecked)

    def update_tournament_options(self):
        is_group_ko = (self.tourn_type_combo.currentText() == "Group Stage + Knockout")
        is_elim_or_ko = (self.tourn_type_combo.currentText() == "Elimination" or is_group_ko)
        for w in [self.tourn_groups_label, self.tourn_groups_spin, self.tourn_qualifiers_label, self.tourn_qualifiers_spin]: w.setVisible(is_group_ko)
        self.tourn_seeding_check.setVisible(is_elim_or_ko)

    def run_tournament(self):
        # Disable run button immediately to prevent double clicks
        self.tourn_run_button.setEnabled(False)
        QApplication.processEvents()

        participants = self.get_selected_participants()
        if len(participants) < 2:
            QMessageBox.warning(self, "Setup Error", "Need >= 2 participants.");
            self.tourn_run_button.setEnabled(True); return # Re-enable button on error

        tourn_type = self.tourn_type_combo.currentText(); rounds_per_game = self.tourn_rounds_spin.value()
        use_seeding = self.tourn_seeding_check.isChecked() and self.tourn_seeding_check.isVisible()
        num_groups = self.tourn_groups_spin.value() if self.tourn_groups_label.isVisible() else 1
        num_qualifiers = self.tourn_qualifiers_spin.value() if self.tourn_qualifiers_label.isVisible() else 1

        # Validation
        if tourn_type == "Group Stage + Knockout":
            if len(participants) < num_groups * 2: QMessageBox.warning(self, "Setup Error", "Too few participants for groups."); self.tourn_run_button.setEnabled(True); return
            if num_qualifiers * num_groups < 2: QMessageBox.warning(self, "Setup Error", "Need >= 2 total qualifiers."); self.tourn_run_button.setEnabled(True); return
            if num_qualifiers < 1: QMessageBox.warning(self, "Setup Error", "Need >= 1 qualifier per group."); self.tourn_run_button.setEnabled(True); return

        tournament_id = f"T-{datetime.datetime.now().strftime('%y%m%d%H%M')}-{str(uuid.uuid4())[:4]}"
        start_time = datetime.datetime.now()
        self.status_bar.showMessage(f"Running {tourn_type} tournament ({tournament_id})..."); QApplication.processEvents()

        # --- Data Storage ---
        tournament_scores = defaultdict(int) # Total score in this tournament
        games_played = defaultdict(int) # Games played in this tournament
        group_stage_scores = defaultdict(int) # For Group+KO ranking
        group_stage_games = defaultdict(int) # For Group+KO ranking
        all_matches_data = [] # Store match details for bulk logging

        if use_seeding:
            participants.sort(key=lambda p_id: STRATEGIES.get(p_id, {}).get('name', 'zzzz')) # Default last if name missing


        final_standings = [] # Use list of dicts directly
        winner_id = None

        # --- Tournament Execution ---
        try:
            if tourn_type == "Round Robin":
                matchups = list(itertools.combinations(participants, 2)); total_matches = len(matchups)
                for i, (p1_id, p2_id) in enumerate(matchups):
                    self.status_bar.showMessage(f"RR Match {i+1}/{total_matches}"); QApplication.processEvents()
                    game = Game(p1_id, p2_id, rounds_per_game, self.global_noise, self.global_forgiveness)
                    s1, s2, h1, h2 = game.run_game()
                    tournament_scores[p1_id] += s1; tournament_scores[p2_id] += s2
                    games_played[p1_id] += 1; games_played[p2_id] += 1
                    self.update_player_stats(p1_id, p2_id, s1, s2) # Update global stats
                    all_matches_data.append({"sim_type": "Tournament Round Robin", "p1_id": p1_id, "p2_id": p2_id, "score1": s1, "score2": s2, "hist1": h1, "hist2": h2})

                # --- Ranking for Round Robin ---
                ranked = sorted(tournament_scores.items(), key=lambda item: item[1], reverse=True)
                winner_id = ranked[0][0] if ranked else None
                final_standings = [{"rank": i + 1, "id": p, "name": STRATEGIES.get(p, {}).get('name', p), "score": sc, "games": games_played.get(p, 0), "avg_score": sc / games_played.get(p, 1) if games_played.get(p, 0) > 0 else 0}
                                   for i, (p, sc) in enumerate(ranked)]


            elif tourn_type == "Elimination":
                current_round_participants = list(participants); round_num = 1
                # Track losers for full ranking based on total score
                eliminated_this_round = []
                full_ranking_data = {p_id: {'score': 0, 'games': 0, 'elim_round': float('inf')} for p_id in participants}

                while len(current_round_participants) > 1:
                    self.status_bar.showMessage(f"Elimination Round {round_num} ({len(current_round_participants)} players)"); QApplication.processEvents()
                    next_round = []; eliminated_this_round = []
                    if len(current_round_participants) % 2 != 0:
                        bye_player = current_round_participants.pop(0); next_round.append(bye_player)
                        print(f"  Bye: {STRATEGIES.get(bye_player, {}).get('name', bye_player)}")

                    matchups = list(zip(current_round_participants[::2], current_round_participants[1::2]))
                    for p1_id, p2_id in matchups:
                        print(f"  Match: {STRATEGIES.get(p1_id,{}).get('name',p1_id)} vs {STRATEGIES.get(p2_id,{}).get('name',p2_id)}"); QApplication.processEvents()
                        game = Game(p1_id, p2_id, rounds_per_game, self.global_noise, self.global_forgiveness); s1, s2, h1, h2 = game.run_game()

                        # Update running totals for final ranking
                        full_ranking_data[p1_id]['score'] += s1; full_ranking_data[p1_id]['games'] += 1
                        full_ranking_data[p2_id]['score'] += s2; full_ranking_data[p2_id]['games'] += 1
                        # Also update global stats
                        self.update_player_stats(p1_id, p2_id, s1, s2)
                        all_matches_data.append({"sim_type": "Tournament Elimination", "p1_id": p1_id, "p2_id": p2_id, "score1": s1, "score2": s2, "hist1": h1, "hist2": h2})

                        winner = p1_id if s1 >= s2 else p2_id # P1 wins ties
                        loser = p2_id if winner == p1_id else p1_id
                        next_round.append(winner)
                        eliminated_this_round.append(loser)
                        full_ranking_data[loser]['elim_round'] = round_num # Mark elimination round for loser
                        print(f"    Winner: {STRATEGIES.get(winner,{}).get('name',winner)}")

                    current_round_participants = next_round
                    round_num += 1

                # --- Ranking for Elimination ---
                winner_id = current_round_participants[0] if current_round_participants else None
                if winner_id:
                    full_ranking_data[winner_id]['elim_round'] = round_num # Winner "eliminated" last

                # Sort primarily by elimination round (higher is better), then by total score
                ranked_ids = sorted(full_ranking_data.keys(), key=lambda p_id: (-full_ranking_data[p_id]['elim_round'], -full_ranking_data[p_id]['score']))

                final_standings = [{"rank": i + 1, "id": p_id, "name": STRATEGIES.get(p_id, {}).get('name', p_id),
                                    "score": full_ranking_data[p_id]['score'], "games": full_ranking_data[p_id]['games'],
                                    "avg_score": full_ranking_data[p_id]['score'] / full_ranking_data[p_id]['games'] if full_ranking_data[p_id]['games'] > 0 else 0}
                                   for i, p_id in enumerate(ranked_ids)]


            elif tourn_type == "Group Stage + Knockout":
                # --- Group Stage ---
                groups = [[] for _ in range(num_groups)]; shuffled = list(participants); random.shuffle(shuffled)
                for i, p_id in enumerate(shuffled): groups[i % num_groups].append(p_id)
                qualifiers = []
                self.status_bar.showMessage("Running Group Stage..."); QApplication.processEvents()

                for g_idx, group in enumerate(groups):
                    if not group: continue # Skip empty group
                    print(f"  Group {g_idx+1} Matches ({[STRATEGIES.get(p, {}).get('name', p) for p in group]}):")
                    matchups = list(itertools.combinations(group, 2))
                    for p1_id, p2_id in matchups:
                         print(f"    {STRATEGIES.get(p1_id,{}).get('name',p1_id)} vs {STRATEGIES.get(p2_id,{}).get('name',p2_id)}"); QApplication.processEvents()
                         game = Game(p1_id, p2_id, rounds_per_game, self.global_noise, self.global_forgiveness); s1, s2, h1, h2 = game.run_game()
                         # Group Stage specific scores
                         group_stage_scores[p1_id] += s1; group_stage_scores[p2_id] += s2
                         group_stage_games[p1_id] += 1; group_stage_games[p2_id] += 1
                         # Also update global stats and total tournament score/games
                         tournament_scores[p1_id] += s1; tournament_scores[p2_id] += s2
                         games_played[p1_id] += 1; games_played[p2_id] += 1
                         self.update_player_stats(p1_id, p2_id, s1, s2)
                         all_matches_data.append({"sim_type": "Tournament Group Stage", "p1_id": p1_id, "p2_id": p2_id, "score1": s1, "score2": s2, "hist1": h1, "hist2": h2})

                    # Determine qualifiers (handle ties simply by score, then name if needed)
                    group_ranked = sorted([p for p in group], key=lambda x: (group_stage_scores.get(x, 0), STRATEGIES.get(x,{}).get('name','')), reverse=True)
                    group_qualifiers = group_ranked[:num_qualifiers]; qualifiers.extend(group_qualifiers)
                    print(f"    Qualifiers: {[STRATEGIES.get(q,{}).get('name', q) for q in group_qualifiers]}")

                # --- Knockout Stage ---
                if len(qualifiers) < 2:
                    winner_id = qualifiers[0] if qualifiers else None
                    QMessageBox.information(self,"Info","Tournament ends after group stage (less than 2 qualifiers).")
                else:
                    if use_seeding: qualifiers.sort(key=lambda p_id: STRATEGIES.get(p_id, {}).get('name', 'zzzz'))
                    print(f"\nKnockout Stage Participants ({len(qualifiers)}): {[STRATEGIES.get(q,{}).get('name',q) for q in qualifiers]}")

                    current_round_participants = qualifiers; round_num = 1
                    while len(current_round_participants) > 1:
                        self.status_bar.showMessage(f"Knockout Round {round_num} ({len(current_round_participants)} players)"); QApplication.processEvents()
                        next_round = []; bye_player = None
                        if len(current_round_participants) % 2 != 0:
                            bye_player = current_round_participants.pop(0); next_round.append(bye_player)
                            print(f"  Bye: {STRATEGIES.get(bye_player, {}).get('name', bye_player)}")

                        matchups = list(zip(current_round_participants[::2], current_round_participants[1::2]))
                        for p1_id, p2_id in matchups:
                            print(f"  Knockout Match: {STRATEGIES.get(p1_id,{}).get('name',p1_id)} vs {STRATEGIES.get(p2_id,{}).get('name',p2_id)}"); QApplication.processEvents()
                            game = Game(p1_id, p2_id, rounds_per_game, self.global_noise, self.global_forgiveness); s1, s2, h1, h2 = game.run_game()
                            # Update total tournament scores and global stats
                            tournament_scores[p1_id] += s1; tournament_scores[p2_id] += s2
                            games_played[p1_id] += 1; games_played[p2_id] += 1
                            self.update_player_stats(p1_id, p2_id, s1, s2)
                            all_matches_data.append({"sim_type": "Tournament Knockout", "p1_id": p1_id, "p2_id": p2_id, "score1": s1, "score2": s2, "hist1": h1, "hist2": h2})
                            winner = p1_id if s1 >= s2 else p2_id # P1 wins ties
                            next_round.append(winner)
                            print(f"    Winner: {STRATEGIES.get(winner,{}).get('name',winner)}")
                        current_round_participants = next_round; round_num += 1
                    winner_id = current_round_participants[0] if current_round_participants else None

                # --- Ranking for Group+KO (based on Group Stage scores) ---
                group_ranked_all = sorted(group_stage_scores.items(), key=lambda item: item[1], reverse=True)
                final_standings = [{"rank": i + 1, "id": p, "name": STRATEGIES.get(p, {}).get('name', p),
                                    "score": sc, "games": group_stage_games.get(p, 0),
                                    "avg_score": sc / group_stage_games.get(p, 1) if group_stage_games.get(p, 0) > 0 else 0}
                                   for i, (p, sc) in enumerate(group_ranked_all) if p in participants] # Filter to original participants


        except Exception as e:
            error_msg = f"Tournament Error: {e}"
            print(error_msg) # Print to console
            import traceback
            traceback.print_exc() # Print full traceback
            self.status_bar.showMessage(error_msg);
            QMessageBox.critical(self, "Error", f"Tournament failed: {e}")
            self.tourn_run_button.setEnabled(True); return # Re-enable button and exit

        # --- Tournament Finished ---

        # Log all matches (moved outside try block)
        try:
            for match_data in all_matches_data:
                 excel_logger.log_match(sim_type=match_data["sim_type"], tournament_id=tournament_id, tournament_type=tourn_type,
                                        p1_id=match_data["p1_id"], p2_id=match_data["p2_id"], rounds=rounds_per_game,
                                        noise_perc=self.global_noise*100, forgiveness_perc=self.global_forgiveness*100,
                                        score1=match_data["score1"], score2=match_data["score2"], history1=match_data["hist1"], history2=match_data["hist2"])
        except Exception as e_log:
             print(f"Error logging matches after tournament completion: {e_log}")

        # Log tournament summary
        winner_name = STRATEGIES.get(winner_id, {}).get('name', 'N/A') if winner_id else 'N/A'
        self.status_bar.showMessage(f"Tournament finished. Winner: {winner_name}. Logging summary.")
        try:
            excel_logger.log_tournament_summary(tournament_id=tournament_id, tournament_type=tourn_type, start_time=start_time, participants_ids=participants,
                                            rounds_per_game=rounds_per_game, noise_perc=self.global_noise*100, forgiveness_perc=self.global_forgiveness*100,
                                            winner_id=winner_id, final_standings=final_standings)
        except Exception as e_log_summary:
             print(f"Error logging tournament summary: {e_log_summary}")


        # Display Leaderboard
        if final_standings:
            # Pass the list of dictionaries directly to the dialog
            LeaderboardDialog(final_standings, tourn_type, self).exec()
        else:
            QMessageBox.information(self, "Tournament Complete", "No standings data to display.")

        self.update_analytics_plot()
        self.tourn_run_button.setEnabled(True) # Re-enable run button


    # --- Custom Strategy Dialog Logic Implementation ---
    def open_define_custom_strategy_dialog(self):
        dialog = CustomStrategyDialog(self)
        if dialog.exec():
            strategy_data = dialog.get_strategy_data()
            if strategy_data:
                 strat_id = strategy_data['id']
                 # Check for duplicate NAME as well as ID
                 existing_id = None
                 existing_name = None
                 for sid, sdata in STRATEGIES.items():
                     if sid == strat_id: existing_id = sid
                     if sdata['name'] == strategy_data['name'] and sid != strat_id: existing_name = sdata['name']

                 if existing_id:
                      if QMessageBox.question(self, "Overwrite?", f"Strategy ID '{strat_id}' exists (from sanitized name). Overwrite?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No) == QMessageBox.StandardButton.No: return
                 elif existing_name:
                      QMessageBox.warning(self, "Name Exists", f"A different strategy already has the name '{existing_name}'. Please choose a unique name.")
                      return


                 try:
                     strategy_func = create_custom_strategy_function(strategy_data['rules'])
                     # Combine data safely using .get for optional fields
                     full_data = {
                        'name': strategy_data['name'], 'func': strategy_func,
                        'desc': strategy_data.get('desc', 'No description.'),
                        'pros_cons': strategy_data.get('pros_cons', 'N/A'),
                        'analogue': strategy_data.get('analogue', 'N/A'),
                        'is_custom': True, 'rules': strategy_data['rules'], 'id': strat_id
                     }
                     STRATEGIES[strat_id] = full_data
                     CUSTOM_STRATEGIES[strat_id] = full_data # Keep CUSTOM_STRATEGIES updated
                     self.update_all_strategy_selectors()
                     index = self.strat_info_combo.findData(strat_id)
                     if index >= 0: self.strat_info_combo.setCurrentIndex(index)
                     self.update_strategy_info_display()
                     self.status_bar.showMessage(f"Custom strategy '{strategy_data['name']}' added/updated. Save recommended.", 4000)
                 except Exception as e_create:
                      QMessageBox.critical(self, "Error", f"Failed to create strategy function: {e_create}")


    def load_and_update_custom_strategies(self):
        # Load function now correctly populates STRATEGIES and CUSTOM_STRATEGIES
        load_custom_strategies()
        self.update_all_strategy_selectors()
        self.update_strategy_info_display()
        count = len(CUSTOM_STRATEGIES)
        self.status_bar.showMessage(f"{count} custom strategies loaded/reloaded from {CUSTOM_STRATEGIES_FILE}.", 4000)


    # --- Sandbox Logic Implementation ---
    def update_sandbox_ui(self):
        # Check if controls exist before accessing them (safety during init)
        if hasattr(self, 'sandbox_p1_combo'):
            self.sandbox_p1_manual = (self.sandbox_p1_combo.currentData() == "manual")
            self.sandbox_p1_manual_group.setVisible(self.sandbox_p1_manual)
        if hasattr(self, 'sandbox_p2_combo'):
            self.sandbox_p2_manual = (self.sandbox_p2_combo.currentData() == "manual")
            self.sandbox_p2_manual_group.setVisible(self.sandbox_p2_manual)
        self.check_sandbox_ready_for_next_round()


    def sandbox_reset(self):
        self.sandbox_game = None; self.sandbox_p1_move = None; self.sandbox_p2_move = None
        p1_id = self.sandbox_p1_combo.currentData(); p2_id = self.sandbox_p2_combo.currentData()
        if not p1_id or not p2_id: QMessageBox.warning(self, "Error", "Select players."); self.sandbox_next_round_button.setEnabled(False); return

        p1_game_id = "manual_p1" if p1_id == "manual" else p1_id
        p2_game_id = "manual_p2" if p2_id == "manual" else p2_id

        try:
            # Fetch names safely using .get
            p1_name = "P1 (Manual)" if p1_id == "manual" else STRATEGIES.get(p1_id, {}).get('name', p1_id)
            p2_name = "P2 (Manual)" if p2_id == "manual" else STRATEGIES.get(p2_id, {}).get('name', p2_id)

            # Use dummy IDs for Game if manual, but correct names for display
            self.sandbox_game = Game(p1_game_id, p2_game_id, MAX_ROUNDS, self.global_noise, self.global_forgiveness)
            self.sandbox_game.strategy1_name = p1_name # Override name in Game object for display
            self.sandbox_game.strategy2_name = p2_name

            self.visualization_widget.clear_data()
            self.visualization_widget.update_data((0, 0, None, None, 0), "", "", p1_name, p2_name, MAX_ROUNDS, status="Ready")
            self.status_bar.showMessage("Sandbox reset. Make moves or click Next Round."); self.sandbox_next_round_button.setEnabled(True)
            self.check_sandbox_ready_for_next_round()
        except Exception as e:
             QMessageBox.critical(self, "Error", f"Failed to reset sandbox: {e}")
             self.sandbox_next_round_button.setEnabled(False)


    def check_sandbox_ready_for_next_round(self):
         # Check if UI elements exist before accessing
         if not hasattr(self, 'sandbox_game') or not hasattr(self, 'sandbox_next_round_button'):
              return # UI not fully initialized yet

         if not self.sandbox_game:
              is_ready = False
              can_p1_move = False
              can_p2_move = False
         else:
             p1_ready = not self.sandbox_p1_manual or self.sandbox_p1_move is not None
             p2_ready = not self.sandbox_p2_manual or self.sandbox_p2_move is not None
             is_ready = p1_ready and p2_ready
             can_p1_move = self.sandbox_p1_manual and self.sandbox_p1_move is None
             can_p2_move = self.sandbox_p2_manual and self.sandbox_p2_move is None

         self.sandbox_next_round_button.setEnabled(is_ready)
         if hasattr(self, 'sandbox_p1_coop_button'):
             self.sandbox_p1_coop_button.setEnabled(can_p1_move)
             self.sandbox_p1_defect_button.setEnabled(can_p1_move)
         if hasattr(self, 'sandbox_p2_coop_button'):
             self.sandbox_p2_coop_button.setEnabled(can_p2_move)
             self.sandbox_p2_defect_button.setEnabled(can_p2_move)


    def sandbox_manual_move(self, player_index, move):
        if not self.sandbox_game: return
        if player_index == 0 and self.sandbox_p1_manual: self.sandbox_p1_move = move; self.status_bar.showMessage(f"P1 chose {move}.")
        elif player_index == 1 and self.sandbox_p2_manual: self.sandbox_p2_move = move; self.status_bar.showMessage(f"P2 chose {move}.")
        self.check_sandbox_ready_for_next_round()


    def sandbox_play_next_round(self):
        if not self.sandbox_game: self.sandbox_reset(); return # Try reset if no game
        p1_id = self.sandbox_p1_combo.currentData(); p2_id = self.sandbox_p2_combo.currentData()
        round_num = len(self.sandbox_game.history1)

        try:
            m1_int = self.sandbox_p1_move if self.sandbox_p1_manual else STRATEGIES[p1_id]['func'](self.sandbox_game.history1, self.sandbox_game.history2, round_num, self.global_forgiveness)
            m2_int = self.sandbox_p2_move if self.sandbox_p2_manual else STRATEGIES[p2_id]['func'](self.sandbox_game.history2, self.sandbox_game.history1, round_num, self.global_forgiveness)

            if m1_int is None or m2_int is None:
                 QMessageBox.warning(self, "Input Needed", "Manual input required."); return

            m1_act = m1_int; m2_act = m2_int
            noise_applied_p1 = False; noise_applied_p2 = False
            if random.random() < self.global_noise: m1_act = COOPERATE if m1_int == DEFECT else DEFECT; noise_applied_p1 = True
            if random.random() < self.global_noise: m2_act = COOPERATE if m2_int == DEFECT else DEFECT; noise_applied_p2 = True

            self.sandbox_game.history1.append(m1_act); self.sandbox_game.history2.append(m2_act)
            payoff1, payoff2 = PAYOFFS[(m1_act, m2_act)]; self.sandbox_game.score1 += payoff1; self.sandbox_game.score2 += payoff2

            state = self.sandbox_game.get_current_state()
            self.visualization_widget.update_data(state, "".join(self.sandbox_game.history1), "".join(self.sandbox_game.history2),
                                                  self.sandbox_game.strategy1_name, self.sandbox_game.strategy2_name, MAX_ROUNDS, status="Running")

            noise_msg = ""
            if noise_applied_p1 and noise_applied_p2: noise_msg = " (Noise P1&P2!)"
            elif noise_applied_p1: noise_msg = " (Noise P1!)"
            elif noise_applied_p2: noise_msg = " (Noise P2!)"

            self.status_bar.showMessage(f"R{round_num+1}: P1->{m1_act}, P2->{m2_act}. Scores: {self.sandbox_game.score1}-{self.sandbox_game.score2}{noise_msg}")
            self.sandbox_p1_move = None; self.sandbox_p2_move = None; self.check_sandbox_ready_for_next_round()

        except KeyError as e:
             QMessageBox.critical(self, "Error", f"Sandbox Error: Unknown strategy ID used: {e}")
             self.sandbox_reset() # Reset on error
        except Exception as e:
             QMessageBox.critical(self, "Error", f"Unexpected Sandbox Error: {e}")
             self.sandbox_reset()


    # --- Analytics Plotting Implementation ---
    def update_analytics_plot(self):
        # Check if canvas exists (might be called early)
        if not hasattr(self, 'analytics_canvas'): return

        plot_type = self.plot_type_combo.currentText(); ax = self.analytics_canvas.axes; ax.clear()
        # Use a copy of PLAYER_STATS for safety
        stats_copy = PLAYER_STATS.copy()
        valid_stats = {k: v for k, v in stats_copy.items() if isinstance(v, dict) and v.get('games_played', 0) > 0 and k in STRATEGIES}

        if not valid_stats:
            ax.text(0.5, 0.5, "No statistics available.", ha='center', va='center', transform=ax.transAxes)
            self.analytics_canvas.draw(); return

        # Sort by name, handling potential missing strategies gracefully
        sorted_ids = sorted(valid_stats.keys(), key=lambda k: STRATEGIES.get(k, {}).get('name', k))
        names = [STRATEGIES.get(k, {}).get('name', k) for k in sorted_ids]
        colors = [('darkblue' if STRATEGIES.get(k, {}).get('is_custom', False) else 'darkgreen') for k in sorted_ids] # Custom=blue, Builtin=green

        try:
            if plot_type == "Total Score":
                vals = [valid_stats[k]['total_score'] for k in sorted_ids]; ax.bar(names, vals, color=colors)
                ax.set_ylabel("Total Score"); ax.set_title("Total Score per Strategy")
            elif plot_type == "Win Rate (%)":
                vals = [(valid_stats[k]['wins'] / valid_stats[k]['games_played']) * 100 for k in sorted_ids]
                ax.bar(names, vals, color=colors); ax.set_ylabel("Win Rate (%)"); ax.set_title("Win Rate per Strategy"); ax.set_ylim(0, 105) # Give margin
            elif plot_type == "Average Score per Game":
                vals = [valid_stats[k]['total_score'] / valid_stats[k]['games_played'] for k in sorted_ids]
                ax.bar(names, vals, color=colors); ax.set_ylabel("Avg Score / Game"); ax.set_title("Average Score per Game")
            elif "Placeholder" in plot_type:
                ax.text(0.5, 0.5, f"{plot_type.replace(' (Placeholder)', '')}\n(Not Implemented)", ha='center', va='center', transform=ax.transAxes)
            else: ax.text(0.5, 0.5, "Select Plot Type", ha='center', va='center', transform=ax.transAxes)

            if "Placeholder" not in plot_type and names:
                 ax.tick_params(axis='x', rotation=70, labelsize=9) # Rotate labels
                 ax.grid(axis='y', linestyle='--', alpha=0.6)

        except Exception as e:
             print(f"Error during plotting: {e}")
             ax.clear()
             ax.text(0.5, 0.5, f"Error plotting '{plot_type}'.", ha='center', va='center', transform=ax.transAxes)


        try: # Add robust tight_layout call
             self.analytics_canvas.fig.tight_layout()
        except Exception as e_layout:
             print(f"Warning: Plot tight_layout failed: {e_layout}")
        self.analytics_canvas.draw()


    # --- Teaching Aids / Scenario Implementation ---
    def load_scenario(self):
        scenario = self.scenario_combo.currentText()
        if scenario == "Select Scenario...": return
        self.select_none_participants(); selected_ids = []
        if scenario == "Axelrod's Participants": selected_ids = ["tit_for_tat", "grudger", "cooperate", "defect", "random", "pavlov", "majority", "suspicious_tft", "tit_for_two_tats", "prober", "generous_tft_10"] # Expanded example set
        elif scenario == "All Cooperate vs All Defect": selected_ids = ["cooperate", "defect"]
        elif scenario == "TFT vs Suspicious TFT": selected_ids = ["tit_for_tat", "suspicious_tft"]
        elif scenario == "All Built-in": selected_ids = list(BUILT_IN_STRATEGIES_META.keys())

        found = 0; missing_names = []
        available_ids = [self.tourn_participants_list.item(i).data(Qt.ItemDataRole.UserRole) for i in range(self.tourn_participants_list.count())]

        # Check requested IDs against available IDs
        for sid in selected_ids:
             if sid not in available_ids:
                  # Try to get the intended name even if missing
                  missing_names.append(BUILT_IN_STRATEGIES_META.get(sid, {}).get('name', sid))

        # Now check the items in the list widget
        for i in range(self.tourn_participants_list.count()):
            item = self.tourn_participants_list.item(i)
            item_id = item.data(Qt.ItemDataRole.UserRole)
            if item_id in selected_ids:
                item.setCheckState(Qt.CheckState.Checked); found += 1

        if missing_names:
             QMessageBox.warning(self, "Warning", f"Scenario '{scenario}' includes strategies currently unavailable: {', '.join(missing_names)}")
        else:
             self.status_bar.showMessage(f"Scenario '{scenario}' loaded. {found} participants selected.", 5000)

        self.tab_widget.setCurrentIndex(1) # Switch to Tournament tab


    # --- Application Closing Implementation ---
    def closeEvent(self, event):
        # Make the excel_logger accessible if it was created globally
        global excel_logger

        reply = QMessageBox.question(self, "Confirm Quit", "Save stats, custom strategies, and log info before exiting?",
                                     QMessageBox.StandardButton.Save | QMessageBox.StandardButton.Discard | QMessageBox.StandardButton.Cancel)
        if reply == QMessageBox.StandardButton.Save:
            print("Saving data...")
            try: save_stats()
            except Exception as e: print(f"Error saving stats: {e}")

            try: save_custom_strategies()
            except Exception as e: print(f"Error saving custom strategies: {e}")

            # Check if excel_logger is the dummy or real one before calling methods on it
            if isinstance(excel_logger, ExcelLogger): # Check if it's the real class
                try:
                     print("Updating Excel strategy info...")
                     excel_logger.update_strategy_info(STRATEGIES)
                     print("Excel strategy info update attempted.")
                except Exception as e: print(f"Error updating strategy log: {e}")
            else:
                 print("Skipping Excel log update as logger initialization failed.")

            print("Exiting.")
            event.accept() # Proceed with closing
        elif reply == QMessageBox.StandardButton.Discard:
            print("Exiting without saving.")
            event.accept() # Proceed with closing without saving
        else:
            print("Quit cancelled.")
            event.ignore() # Don't close the application


# Block 8: Main Execution Block

if __name__ == '__main__':
    app = QApplication(sys.argv)
    excel_logger = None # Initialize logger variable

    # --- Pre-computation / Loading ---
    print("Loading stats...")
    load_stats() # Load player stats into PLAYER_STATS global
    print("Loading custom strategies...")
    load_custom_strategies() # Load custom strategies and update STRATEGIES global
    print(f"Strategies loaded: {list(STRATEGIES.keys())}")

    if not STRATEGIES:
        QMessageBox.critical(None, "Startup Error", "No strategies (built-in or custom) could be loaded. Cannot start.")
        sys.exit(1)

    # --- Initialize Excel Logger ---
    try:
        print("Initializing Excel Logger...")
        # Ensure logger is created AFTER strategies might be loaded (for P1/P2 Name lookup if needed)
        excel_logger = ExcelLogger()
        print("Excel Logger Initialized.")
    except Exception as e:
         # Use dummy logger if Excel init fails, prevents crashes on logging calls
         print(f"CRITICAL: Excel Logger initialization failed: {e}. Logging disabled.")
         # Define a dummy class locally for this scope
         class DummyLogger:
             def __getattr__(self, name):
                 # Return a function that does nothing when any method is called
                 def _dummy_method(*args, **kwargs):
                     # print(f"DummyLogger: Call to {name} ignored.") # Optional: for debugging
                     pass
                 return _dummy_method
         excel_logger = DummyLogger()
         QMessageBox.critical(None,"Excel Error", f"Failed to initialize Excel log '{LOG_FILE}': {e}\nLogging is disabled for this session.")


    # --- Create and Show Main Window ---
    print("Creating main window...")
    main_window = IPDSimulatorV6()
    print("Showing main window...")
    main_window.show()
    print("Starting application event loop...")
    sys.exit(app.exec())
