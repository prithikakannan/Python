import random
from datetime import datetime
import polars as pl
import os

class RPSGame:
    def __init__(self):
        self.excel_path = os.path.join(os.path.dirname(__file__), 'game_history.xlsx')
        
        # Initialize or load game history
        if os.path.exists(self.excel_path):
            self.load_history_from_excel()
        else:
            self.game_history = []
            self.move_counts = {'rock': 0, 'paper': 0, 'scissors': 0}
            self.win_rates = [0]  # Initialize with 0
            self.wins = 0
            self.total_games = 0
            self.init_excel_file()
    
    def init_excel_file(self):
        # Create initial Excel file with empty dataframe
        history_df = pl.DataFrame(
            schema={
                'datetime': pl.Utf8,
                'player': pl.Utf8,
                'computer': pl.Utf8,
                'result': pl.Utf8
            }
        )
        
        stats_df = pl.DataFrame({
            'total_games': [0],
            'wins': [0],
            'rock_count': [0],
            'paper_count': [0],
            'scissors_count': [0],
            'win_rate': [0.0]
        })
        
        # Use pandas to write to Excel (polars doesn't support Excel writing directly)
        import pandas as pd
        with pd.ExcelWriter(self.excel_path) as writer:
            history_df.to_pandas().to_excel(writer, sheet_name='History', index=False)
            stats_df.to_pandas().to_excel(writer, sheet_name='Stats', index=False)
    
    def load_history_from_excel(self):
        # Use pandas to read Excel then convert to polars
        import pandas as pd
        
        # Load game history from Excel
        history_df = pl.from_pandas(pd.read_excel(self.excel_path, sheet_name='History'))
        stats_df = pl.from_pandas(pd.read_excel(self.excel_path, sheet_name='Stats'))
        
        # Convert DataFrame to list of dictionaries
        self.game_history = history_df.to_dicts()
        
        # Load statistics
        if not stats_df.is_empty():
            self.total_games = stats_df.item(0, 'total_games')
            self.wins = stats_df.item(0, 'wins')
            self.move_counts = {
                'rock': stats_df.item(0, 'rock_count'),
                'paper': stats_df.item(0, 'paper_count'),
                'scissors': stats_df.item(0, 'scissors_count')
            }
            
            # Calculate win rates array
            if self.total_games > 0:
                win_rate = (self.wins / self.total_games) * 100
            else:
                win_rate = 0
            
            self.win_rates = [win_rate]  # Just use the current rate for initial trend
            
            # Generate trend data based on history
            self._generate_trend_from_history()
    
    def _generate_trend_from_history(self):
        # Reset win rates
        self.win_rates = [0]
        
        # Track running total for win calculation
        running_wins = 0
        running_games = 0
        
        for game in self.game_history:
            running_games += 1
            if game['result'] == 'wins':
                running_wins += 1
            
            # Calculate win rate at this point
            if running_games > 0:
                win_rate = (running_wins / running_games) * 100
            else:
                win_rate = 0
                
            self.win_rates.append(win_rate)
    
    def save_to_excel(self):
        # Convert game history to DataFrame
        history_df = pl.DataFrame(self.game_history)
        
        # Create stats DataFrame
        stats_df = pl.DataFrame({
            'total_games': [self.total_games],
            'wins': [self.wins],
            'rock_count': [self.move_counts['rock']],
            'paper_count': [self.move_counts['paper']],
            'scissors_count': [self.move_counts['scissors']],
            'win_rate': [(self.wins / self.total_games * 100) if self.total_games > 0 else 0]
        })
        
        # Use pandas to write to Excel
        import pandas as pd
        with pd.ExcelWriter(self.excel_path) as writer:
            # Convert polars to pandas before writing to Excel
            if not history_df.is_empty():
                history_df.to_pandas().to_excel(writer, sheet_name='History', index=False)
            else:
                pd.DataFrame(columns=['datetime', 'player', 'computer', 'result']).to_excel(
                    writer, sheet_name='History', index=False
                )
            stats_df.to_pandas().to_excel(writer, sheet_name='Stats', index=False)

    def play(self, player_choice):
        choices = ['rock', 'paper', 'scissors']
        computer_choice = random.choice(choices)
        
        # Update move counts
        self.move_counts[player_choice.lower()] += 1
        
        # Determine winner
        if player_choice == computer_choice:
            result = 'ties'
        elif (
            (player_choice == 'rock' and computer_choice == 'scissors') or
            (player_choice == 'paper' and computer_choice == 'rock') or
            (player_choice == 'scissors' and computer_choice == 'paper')
        ):
            result = 'wins'
            self.wins += 1
        else:
            result = 'losses'
        
        self.total_games += 1
        
        # Calculate win rate
        win_rate = (self.wins / self.total_games) * 100 if self.total_games > 0 else 0
        self.win_rates.append(win_rate)
        
        # Record game in history
        new_game = {
            'datetime': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'player': player_choice,
            'computer': computer_choice,
            'result': result
        }
        self.game_history.append(new_game)
        
        # Save to Excel after each game
        self.save_to_excel()
        
        return computer_choice, result

    def get_stats(self):
        win_rate = f"{(self.wins / self.total_games * 100):.1f}%" if self.total_games > 0 else "0.0%"
        return {
            'total_games': self.total_games,
            'wins': self.wins,
            'win_rate': win_rate
        }

    def get_move_distribution(self):
        return {k.capitalize(): v for k, v in self.move_counts.items()}

    def get_winrate_trend(self):
        return self.win_rates
