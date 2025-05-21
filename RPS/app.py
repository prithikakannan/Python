import customtkinter as ctk
from datetime import datetime
from data_utils import create_charts
from game_utils import RPSGame
import os

# Force dark mode only
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Create the main application window
app = ctk.CTk()
app.title("Rock Paper Scissors Game")

# Set the window size
app.geometry("1100x700")

# Center the window on the screen
screen_width = app.winfo_screenwidth()
screen_height = app.winfo_screenheight()
x = (screen_width - 1100) // 2
y = (screen_height - 700) // 2
app.geometry(f"1100x700+{x}+{y}")

# Define color scheme (dark mode only)
COLORS = {
    'primary': "#8B5CF6",       # Purple main color
    'secondary': "#2D2D44",     # Dark background
    'secondary_alt': "#1E1E2E", # Slightly darker background
    'accent': "#34D399",        # Green accent
    'warning': "#FBBF24",       # Amber warning
    'danger': "#F87171",        # Red danger
    'text': "#F3F4F6",          # Light text
    'text_secondary': "#D1D5DB" # Gray text
}

# Configure grid layout (2x1)
app.grid_rowconfigure(0, weight=1)
app.grid_columnconfigure(1, weight=1)

# Create frames dictionary to store different views
frames = {}

# Create sidebar frame with fixed width
sidebar_frame = ctk.CTkFrame(app, width=240, corner_radius=0, fg_color=COLORS['secondary'])
sidebar_frame.grid(row=0, column=0, sticky="nsew")
sidebar_frame.grid_propagate(False)  # Prevent resizing

# Configure sidebar grid
sidebar_frame.grid_rowconfigure((0, 1, 2, 3, 4, 5, 6, 7), weight=0)  # Navigation items
sidebar_frame.grid_rowconfigure(8, weight=1)  # Spacer
sidebar_frame.grid_rowconfigure(9, weight=0)  # Version at bottom
sidebar_frame.grid_columnconfigure(0, weight=1)  # Center items

# Create main content frame
main_frame = ctk.CTkFrame(app, fg_color="transparent")
main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
main_frame.grid_columnconfigure(0, weight=1)
main_frame.grid_rowconfigure(0, weight=1)

# Add app title to sidebar
logo_label = ctk.CTkLabel(sidebar_frame, text="RPS Game", 
                         font=ctk.CTkFont(family="Helvetica", size=28, weight="bold"),
                         text_color=COLORS['primary'])
logo_label.grid(row=0, column=0, padx=20, pady=(40, 10))

# Add subtitle
subtitle_label = ctk.CTkLabel(sidebar_frame, text="Rock Paper Scissors", 
                             font=ctk.CTkFont(family="Helvetica", size=14),
                             text_color=COLORS['text_secondary'])
subtitle_label.grid(row=1, column=0, padx=20, pady=(0, 30))

# Initialize game
game = RPSGame()

# Navigation button styling
button_style = {
    "height": 45, 
    "corner_radius": 10, 
    "hover_color": COLORS['primary'],
    "fg_color": "transparent",
    "text_color": COLORS['text'],
    "anchor": "w",
    "font": ctk.CTkFont(family="Roboto", size=16)
}

nav_buttons = [
    ("🏠  Dashboard", "dashboard"),
    ("🎮  Play Game", "game"),
    ("📊  Statistics", "history")
]

def show_frame(frame_name):
    for frame in frames.values():
        frame.grid_remove()
    if frame_name == 'dashboard':
        update_stats_display()
    elif frame_name == 'history':
        update_history_display()
    elif frame_name == 'achievements':
        update_achievements_display()
    elif frame_name == 'settings':
        update_settings_display()
    elif frame_name == 'help':
        update_help_display()
    frames[frame_name].grid(row=0, column=0, sticky="nsew")

# Create and configure navigation buttons
for i, (text, frame_name) in enumerate(nav_buttons, 2):
    btn = ctk.CTkButton(sidebar_frame, text=text, command=lambda n=frame_name: show_frame(n), **button_style)
    btn.grid(row=i, column=0, padx=20, pady=10, sticky="ew")

# Add version at the bottom
version_label = ctk.CTkLabel(sidebar_frame, 
                           text="v1.1.0",
                           font=ctk.CTkFont(family="Arial", size=12),
                           text_color=COLORS['text_secondary'])
version_label.grid(row=9, column=0, padx=20, pady=(0, 20), sticky="s")

# Create main dashboard frame - change to scrollable
dashboard_frame = ctk.CTkScrollableFrame(main_frame, fg_color="transparent")
frames['dashboard'] = dashboard_frame
dashboard_frame.grid_columnconfigure(0, weight=1)

# Create other frames
for frame_name in ['game', 'history', 'achievements', 'settings', 'help']:
    frame = ctk.CTkFrame(main_frame, fg_color="transparent")
    frames[frame_name] = frame
    frame.grid_columnconfigure(0, weight=1)

# Dashboard content
# Header with date and time-based greeting
header_frame = ctk.CTkFrame(dashboard_frame, fg_color=COLORS['secondary'], corner_radius=15)
header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 20))
header_frame.grid_columnconfigure(1, weight=1)

# Get time-based greeting
current_hour = datetime.now().hour
if 5 <= current_hour < 12:
    greeting = "Good Morning"
elif 12 <= current_hour < 18:
    greeting = "Good Afternoon"
else:
    greeting = "Good Evening"

welcome_label = ctk.CTkLabel(header_frame, 
                           text=f"🏆 {greeting}, Prithika!",
                           font=ctk.CTkFont(family="Helvetica", size=24, weight="bold"),
                           text_color=COLORS['text'])
welcome_label.grid(row=0, column=0, padx=20, pady=20, sticky="w")

current_time = datetime.now().strftime("%B %d, %Y")
time_label = ctk.CTkLabel(header_frame, text=current_time, 
                         font=ctk.CTkFont(family="Arial", size=15),
                         text_color=COLORS['text_secondary'])
time_label.grid(row=0, column=1, padx=20, pady=20, sticky="e")

# Statistics cards frame
stats_frame = ctk.CTkFrame(dashboard_frame, fg_color="transparent")
stats_frame.grid(row=1, column=0, sticky="ew", pady=(0, 20))
stats_frame.grid_columnconfigure((0,1,2), weight=1)

# Add progress section
progress_frame = ctk.CTkFrame(dashboard_frame, fg_color=COLORS['secondary'], corner_radius=15)
progress_frame.grid(row=2, column=0, sticky="ew", pady=(0, 20))
progress_frame.grid_columnconfigure(0, weight=1)

progress_title = ctk.CTkLabel(progress_frame, 
                             text="Win Progress",
                             font=ctk.CTkFont(family="Helvetica", size=18, weight="bold"),
                             text_color=COLORS['text'])
progress_title.grid(row=0, column=0, padx=20, pady=(15, 5), sticky="w")

# Target explanation
target_text = "Target: 10 wins"
target_label = ctk.CTkLabel(progress_frame,
                           text=target_text,
                           font=ctk.CTkFont(family="Arial", size=14),
                           text_color=COLORS['text_secondary'])
target_label.grid(row=1, column=0, padx=20, pady=(0, 5), sticky="w")

# Progress bar
progress_container = ctk.CTkFrame(progress_frame, fg_color="transparent")
progress_container.grid(row=2, column=0, padx=20, pady=(0, 15), sticky="ew")
progress_container.grid_columnconfigure(0, weight=1)

# Progress bar will be updated in update_stats_display function
progress_bar = ctk.CTkProgressBar(progress_container, height=15, corner_radius=5)
progress_bar.grid(row=0, column=0, sticky="ew", pady=5)
progress_bar.configure(fg_color=COLORS['secondary_alt'], progress_color=COLORS['accent'])

progress_text = ctk.CTkLabel(progress_container,
                            text="0/10 wins (0%)",
                            font=ctk.CTkFont(family="Arial", size=12),
                            text_color=COLORS['text_secondary'])
progress_text.grid(row=1, column=0, sticky="e", pady=(0, 5))

# Analytics charts
charts_frame = ctk.CTkFrame(dashboard_frame, fg_color=COLORS['secondary'], corner_radius=15)
charts_frame.grid(row=3, column=0, sticky="nsew", pady=(0, 20))

# Create and add charts title
charts_title = ctk.CTkLabel(charts_frame, 
                           text="Game Performance Analysis",
                           font=ctk.CTkFont(family="Helvetica", size=18, weight="bold"),
                           text_color=COLORS['text'])
charts_title.pack(anchor="w", padx=20, pady=(15, 0))

# Create and add charts
charts_widget = create_charts(charts_frame)
charts_widget.pack(fill="both", expand=True, padx=10, pady=10)

# Tips and tricks section
tips_frame = ctk.CTkFrame(dashboard_frame, fg_color=COLORS['secondary'], corner_radius=15)
tips_frame.grid(row=4, column=0, sticky="ew", pady=(0, 20))

tips_title = ctk.CTkLabel(tips_frame, 
                         text="💡 Tips & Tricks",
                         font=ctk.CTkFont(family="Helvetica", size=18, weight="bold"),
                         text_color=COLORS['text'])
tips_title.grid(row=0, column=0, padx=20, pady=(15, 10), sticky="w")

# Tips content
tips = [
    "Rock is statistically thrown most often by beginners",
    "Watch for patterns in your opponent's moves",
    "After losing with rock, players often switch to paper",
    "Try to maintain unpredictable play patterns"
]

tips_container = ctk.CTkFrame(tips_frame, fg_color="transparent")
tips_container.grid(row=1, column=0, padx=20, pady=(0, 15), sticky="ew")

for i, tip in enumerate(tips):
    tip_label = ctk.CTkLabel(tips_container,
                            text=f"• {tip}",
                            font=ctk.CTkFont(family="Arial", size=13),
                            text_color=COLORS['text'])
    tip_label.grid(row=i, column=0, padx=5, pady=3, sticky="w")

# Recent games frame
recent_games_frame = ctk.CTkFrame(dashboard_frame, fg_color=COLORS['secondary'], corner_radius=15)
recent_games_frame.grid(row=5, column=0, sticky="ew", pady=(0, 20))

recent_games_title = ctk.CTkLabel(recent_games_frame, 
                                 text="Recent Games",
                                 font=ctk.CTkFont(family="Helvetica", size=18, weight="bold"),
                                 text_color=COLORS['text'])
recent_games_title.pack(anchor="w", padx=20, pady=(15, 10))

# Create stats card
def create_stat_card(parent, icon, value, title, color):
    card = ctk.CTkFrame(parent, fg_color=COLORS['secondary'], corner_radius=15)
    
    icon_label = ctk.CTkLabel(card, text=icon, 
                             font=ctk.CTkFont(family="Segoe UI", size=28),
                             text_color=color)
    icon_label.grid(row=0, column=0, padx=(20, 10), pady=(15, 5))
    
    value_label = ctk.CTkLabel(card, text=value, 
                              font=ctk.CTkFont(family="Helvetica", size=24, weight="bold"),
                              text_color=COLORS['text'])
    value_label.grid(row=0, column=1, padx=10, pady=(15, 5), sticky="w")
    
    title_label = ctk.CTkLabel(card, text=title, 
                              font=ctk.CTkFont(family="Arial", size=14),
                              text_color=COLORS['text_secondary'])
    title_label.grid(row=1, column=0, columnspan=2, padx=20, pady=(0, 15), sticky="w")
    
    return card, value_label

# Create recent games list
def update_recent_games():
    # Clear existing widgets
    for widget in recent_games_frame.winfo_children():
        if widget != recent_games_title:
            widget.destroy()
    
    # Create scrollable frame for recent games
    recent_scroll = ctk.CTkScrollableFrame(recent_games_frame, height=150, fg_color="transparent")
    recent_scroll.pack(fill="both", expand=True, padx=10, pady=(0, 10))
    
    # Get last 5 games from history
    games = game.game_history[-5:] if len(game.game_history) > 0 else []
    
    if not games:
        no_games = ctk.CTkLabel(recent_scroll,
                               text="No games played yet. Start playing!",
                               font=ctk.CTkFont(family="Arial", size=14),
                               text_color=COLORS['text_secondary'])
        no_games.pack(pady=20)
        return
    
    # Display games in reverse order (most recent first)
    for i, record in enumerate(reversed(games)):
        # Create game entry frame
        entry_frame = ctk.CTkFrame(recent_scroll, fg_color=COLORS['secondary_alt'], corner_radius=8)
        entry_frame.pack(fill="x", padx=5, pady=5)
        
        # Format date
        date_str = record['datetime']
        
        # Result icon and color
        if record['result'] == 'wins':
            result_icon = "🏆"
            result_color = COLORS['accent']
        elif record['result'] == 'losses':
            result_icon = "❌"
            result_color = COLORS['danger']
        else:
            result_icon = "🤝"
            result_color = COLORS['warning']
        
        # Game result label
        result_label = ctk.CTkLabel(entry_frame, 
                                   text=f"{result_icon}",
                                   font=ctk.CTkFont(size=20),
                                   text_color=result_color)
        result_label.grid(row=0, column=0, padx=10, pady=10)
        
        # Game details
        details_text = f"You: {record['player'].upper()} vs PC: {record['computer'].upper()}"
        details_label = ctk.CTkLabel(entry_frame, 
                                    text=details_text,
                                    font=ctk.CTkFont(size=14),
                                    text_color=COLORS['text'])
        details_label.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        
        # Date label
        date_label = ctk.CTkLabel(entry_frame, 
                                 text=date_str,
                                 font=ctk.CTkFont(size=12),
                                 text_color=COLORS['text_secondary'])
        date_label.grid(row=0, column=2, padx=10, pady=10, sticky="e")
        
        # Configure grid weights
        entry_frame.grid_columnconfigure(1, weight=1)

# Update dashboard stats
def update_stats_display():
    # Clear existing cards
    for widget in stats_frame.winfo_children():
        widget.destroy()
    
    # Get game stats
    stats = game.get_stats()
    
    # Create stat cards
    cards_data = [
        ("🎮", str(stats['total_games']), "Total Games", COLORS['primary']),
        ("🏆", str(stats['wins']), "Wins", COLORS['accent']),
        ("📈", stats['win_rate'], "Win Rate", COLORS['warning'])
    ]
    
    # Add cards to frame
    for idx, (icon, value, title, color) in enumerate(cards_data):
        card, _ = create_stat_card(stats_frame, icon, value, title, color)
        card.grid(row=0, column=idx, padx=10, pady=5, sticky="ew")
    
    # Update progress bar
    wins = stats['wins']
    target = 10
    progress = min(1.0, wins / target) if target > 0 else 0
    progress_bar.set(progress)
    progress_percent = int(progress * 100)
    progress_text.configure(text=f"{wins}/{target} wins ({progress_percent}%)")
    
    # Update recent games
    update_recent_games()
    
    # Update charts
    analytics_data = {
        'moves': game.get_move_distribution(),
        'trend': game.get_winrate_trend()
    }
    charts_widget.update_charts(analytics_data)

# History/stats display
def update_history_display():
    history_frame = frames['history']
    
    # Clear existing content
    for widget in history_frame.winfo_children():
        widget.destroy()
    
    # Create header
    header = ctk.CTkFrame(history_frame, fg_color=COLORS['secondary'], corner_radius=15)
    header.grid(row=0, column=0, sticky="ew", pady=(0, 20))
    
    title = ctk.CTkLabel(header, text="Game History & Statistics",
                        font=ctk.CTkFont(family="Helvetica", size=24, weight="bold"),
                        text_color=COLORS['text'])
    title.grid(row=0, column=0, padx=20, pady=20, sticky="w")
    
    # Create statistics panels
    stats_panel = ctk.CTkFrame(history_frame, fg_color=COLORS['secondary'], corner_radius=15)
    stats_panel.grid(row=1, column=0, sticky="ew", pady=(0, 20))
    stats_panel.grid_columnconfigure((0,1,2), weight=1)
    
    # Get stats
    stats = game.get_stats()
    
    # Add detailed statistics
    stats_title = ctk.CTkLabel(stats_panel, 
                              text="Detailed Statistics",
                              font=ctk.CTkFont(family="Helvetica", size=18, weight="bold"),
                              text_color=COLORS['text'])
    stats_title.grid(row=0, column=0, columnspan=3, padx=20, pady=(15, 15), sticky="w")
    
    # Calculate losses and ties
    losses = stats['total_games'] - stats['wins']
    ties = sum(1 for g in game.game_history if g['result'] == 'ties')
    
    # Create detailed stat cards
    cards_data = [
        ("🎮", str(stats['total_games']), "Total Games", COLORS['primary']),
        ("🏆", str(stats['wins']), "Wins", COLORS['accent']),
        ("📉", str(losses), "Losses", COLORS['danger']),
        ("🤝", str(ties), "Ties", COLORS['warning']),
        ("📊", stats['win_rate'], "Win Rate", COLORS['primary']),
        ("⏱️", str(len(game.game_history)), "Games Recorded", COLORS['accent'])
    ]
    
    # Add cards to frame, 3 per row
    for idx, (icon, value, title, color) in enumerate(cards_data):
        row = idx // 3 + 1
        col = idx % 3
        card, _ = create_stat_card(stats_panel, icon, value, title, color)
        card.grid(row=row, column=col, padx=10, pady=10, sticky="ew")
    
    # Create scrollable frame for game history
    history_container = ctk.CTkFrame(history_frame, fg_color=COLORS['secondary'], corner_radius=15)
    history_container.grid(row=2, column=0, sticky="nsew", pady=(0, 20))
    history_container.grid_rowconfigure(1, weight=1)
    history_container.grid_columnconfigure(0, weight=1)
    
    history_title = ctk.CTkLabel(history_container, 
                                text="Complete Game History",
                                font=ctk.CTkFont(family="Helvetica", size=18, weight="bold"),
                                text_color=COLORS['text'])
    history_title.grid(row=0, column=0, padx=20, pady=(15, 10), sticky="w")
    
    # Create scrollable frame for the history
    scroll_frame = ctk.CTkScrollableFrame(history_container, fg_color="transparent")
    scroll_frame.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nsew")
    scroll_frame.columnconfigure(0, weight=1)
    
    # Add column headers
    headers = ["Date & Time", "Player Move", "Computer Move", "Result"]
    header_frame = ctk.CTkFrame(scroll_frame, fg_color=COLORS['primary'], corner_radius=0)
    header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 5))
    
    for i, header in enumerate(headers):
        header_label = ctk.CTkLabel(header_frame, 
                                   text=header,
                                   font=ctk.CTkFont(weight="bold"),
                                   text_color=COLORS['text'])
        header_label.grid(row=0, column=i, padx=15, pady=10, sticky="w")
    
    # Configure column weights
    header_frame.columnconfigure(0, weight=2)
    header_frame.columnconfigure(1, weight=1)
    header_frame.columnconfigure(2, weight=1)
    header_frame.columnconfigure(3, weight=1)
    
    # If no games
    if not game.game_history:
        no_games = ctk.CTkLabel(scroll_frame,
                               text="No games played yet. Start playing!",
                               font=ctk.CTkFont(family="Arial", size=14),
                               text_color=COLORS['text_secondary'])
        no_games.grid(row=1, column=0, pady=20)
    else:
        # Add history entries
        for i, record in enumerate(reversed(game.game_history), 1):
            # Alternate row colors
            bg_color = COLORS['secondary_alt'] if i % 2 == 0 else "transparent"
            
            row_frame = ctk.CTkFrame(scroll_frame, fg_color=bg_color, corner_radius=0, height=40)
            row_frame.grid(row=i, column=0, sticky="ew")
            
            # Configure column weights
            row_frame.columnconfigure(0, weight=2)
            row_frame.columnconfigure(1, weight=1)
            row_frame.columnconfigure(2, weight=1)
            row_frame.columnconfigure(3, weight=1)
            
            # Date
            date_label = ctk.CTkLabel(row_frame, 
                                     text=record['datetime'],
                                     font=ctk.CTkFont(size=13),
                                     text_color=COLORS['text'])
            date_label.grid(row=0, column=0, padx=15, pady=10, sticky="w")
            
            # Player move
            player_label = ctk.CTkLabel(row_frame, 
                                       text=record['player'].title(),
                                       font=ctk.CTkFont(size=13),
                                       text_color=COLORS['text'])
            player_label.grid(row=0, column=1, padx=15, pady=10, sticky="w")
            
            # Computer move
            computer_label = ctk.CTkLabel(row_frame, 
                                         text=record['computer'].title(),
                                         font=ctk.CTkFont(size=13),
                                         text_color=COLORS['text'])
            computer_label.grid(row=0, column=2, padx=15, pady=10, sticky="w")
            
            # Result with colored icon
            if record['result'] == 'wins':
                result_text = "Win 🏆"
                result_color = COLORS['accent']
            elif record['result'] == 'losses':
                result_text = "Loss ❌"
                result_color = COLORS['danger']
            else:
                result_text = "Tie 🤝"
                result_color = COLORS['warning']
                
            result_label = ctk.CTkLabel(row_frame, 
                                       text=result_text,
                                       font=ctk.CTkFont(size=13),
                                       text_color=result_color)
            result_label.grid(row=0, column=3, padx=15, pady=10, sticky="w")

# Create game interface
def create_game_interface():
    game_frame = frames['game']
    
    # Clear existing content
    for widget in game_frame.winfo_children():
        widget.destroy()
    
    # Create header
    header = ctk.CTkFrame(game_frame, fg_color=COLORS['secondary'], corner_radius=15)
    header.grid(row=0, column=0, sticky="ew", pady=(0, 20))
    
    title = ctk.CTkLabel(header, text="Play Rock Paper Scissors",
                        font=ctk.CTkFont(family="Helvetica", size=24, weight="bold"),
                        text_color=COLORS['text'])
    title.grid(row=0, column=0, padx=20, pady=20, sticky="w")
    
    # Game container with custom styling
    game_container = ctk.CTkFrame(game_frame, fg_color=COLORS['secondary'], corner_radius=15)
    game_container.grid(row=1, column=0, padx=0, pady=0, sticky="nsew")
    game_container.grid_rowconfigure(3, weight=1)
    game_container.grid_columnconfigure(0, weight=1)
    
    # Game title
    game_title = ctk.CTkLabel(game_container, 
                            text="✨ Choose Your Move! ✨",
                            font=ctk.CTkFont(family="Helvetica", size=28, weight="bold"),
                            text_color=COLORS['primary'])
    game_title.grid(row=0, column=0, pady=(30, 10))
    
    # Score display
    score_frame = ctk.CTkFrame(game_container, fg_color=COLORS['secondary_alt'], corner_radius=10)
    score_frame.grid(row=1, column=0, pady=20, padx=40, sticky="ew")
    
    stats = game.get_stats()
    score_text = f"Games: {stats['total_games']} | Wins: {stats['wins']} | Win Rate: {stats['win_rate']}"
    score_label = ctk.CTkLabel(score_frame, 
                              text=score_text,
                              font=ctk.CTkFont(family="Arial", size=16, weight="bold"),
                              text_color=COLORS['text'])
    score_label.pack(pady=15)
    
    # Result display
    result_frame = ctk.CTkFrame(game_container, 
                               fg_color=COLORS['secondary_alt'],
                               corner_radius=15)
    result_frame.grid(row=2, column=0, pady=20, padx=40, sticky="ew")
    
    result_label = ctk.CTkLabel(result_frame, 
                               text="Make your choice below!",
                               font=ctk.CTkFont(family="Arial", size=20, weight="bold"),
                               text_color=COLORS['text'])
    result_label.pack(pady=30)
    
    # Game logic
    def play_game(choice):
        computer_choice, result = game.play(choice)
        
        # Determine result styling
        if result == 'wins':
            result_color = COLORS['accent']
            result_emoji = "🎉"
            result_text = "You Win!"
        elif result == 'losses':
            result_color = COLORS['danger']
            result_emoji = "💔"
            result_text = "You Lose!"
        else:
            result_color = COLORS['warning']
            result_emoji = "🤝"
            result_text = "It's a Tie!"
            
        # Update result display
        display_text = f"{result_emoji} {result_text} {result_emoji}\n\n"
        display_text += f"You chose {choice.upper()}\nComputer chose {computer_choice.upper()}"
        
        result_label.configure(text=display_text, text_color=result_color)
        
        # Update score
        stats = game.get_stats()
        score_text = f"Games: {stats['total_games']} | Wins: {stats['wins']} | Win Rate: {stats['win_rate']}"
        score_label.configure(text=score_text)
        
        # Update dashboard if it's visible
        update_stats_display()
        update_recent_games()
    
    # Choice buttons
    button_frame = ctk.CTkFrame(game_container, fg_color="transparent")
    button_frame.grid(row=3, column=0, pady=30)
    
    button_style = {
        "width": 160,
        "height": 160,
        "corner_radius": 15,
        "border_width": 2,
        "border_color": COLORS['primary'],
        "hover_color": COLORS['primary'],
        "fg_color": COLORS['secondary_alt'],
        "text_color": COLORS['text'],
        "font": ctk.CTkFont(family="Arial", size=24, weight="bold")
    }
    
    choices = [
        ("🪨", "ROCK", "rock"),
        ("📄", "PAPER", "paper"),
        ("✂️", "SCISSORS", "scissors")
    ]
    
    for i, (emoji, text, choice) in enumerate(choices):
        # Create choice frame
        choice_frame = ctk.CTkFrame(button_frame, fg_color="transparent")
        choice_frame.grid(row=0, column=i, padx=15)
        
        # Create button
        btn = ctk.CTkButton(choice_frame, 
                           text=emoji,
                           command=lambda c=choice: play_game(c),
                           **button_style)
        btn.grid(row=0, column=0, padx=0, pady=0)
        
        # Add label
        label = ctk.CTkLabel(choice_frame,
                            text=text,
                            font=ctk.CTkFont(family="Arial", size=16, weight="bold"),
                            text_color=COLORS['text'])
        label.grid(row=1, column=0, pady=(10, 0))
        
        # Add hover effects
        def on_enter(e, button=btn):
            button.configure(
                fg_color=COLORS['primary'],
                text_color=COLORS['text']
            )
        
        def on_leave(e, button=btn):
            button.configure(
                fg_color=COLORS['secondary_alt'],
                text_color=COLORS['text']
            )
        
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)

# Initialize game interface
create_game_interface()

# Show frame initially
show_frame('dashboard')

# Run the application
if __name__ == "__main__":
    app.mainloop()

# Achievements display
def update_achievements_display():
    achievements_frame = frames['achievements']
    
    # Clear existing content
    for widget in achievements_frame.winfo_children():
        widget.destroy()
    
    # Create header
    header = ctk.CTkFrame(achievements_frame, fg_color=COLORS['secondary'], corner_radius=15)
    header.grid(row=0, column=0, sticky="ew", pady=(0, 20))
    
    title = ctk.CTkLabel(header, text="Game Achievements",
                       font=ctk.CTkFont(family="Helvetica", size=24, weight="bold"),
                       text_color=COLORS['text'])
    title.grid(row=0, column=0, padx=20, pady=20, sticky="w")
    
    # Create achievements container
    achievements_container = ctk.CTkFrame(achievements_frame, fg_color=COLORS['secondary'], corner_radius=15)
    achievements_container.grid(row=1, column=0, sticky="ew", pady=(0, 20))
    
    # Add achievements title
    achievements_title = ctk.CTkLabel(achievements_container, 
                                    text="Your Progress",
                                    font=ctk.CTkFont(family="Helvetica", size=18, weight="bold"),
                                    text_color=COLORS['text'])
    achievements_title.grid(row=0, column=0, padx=20, pady=(15, 15), sticky="w")
    
    # Get stats for achievements
    stats = game.get_stats()
    total_games = stats['total_games']
    total_wins = stats['wins']
    
    # Define achievements
    achievements_list = [
        {
            "title": "Beginner",
            "description": "Play your first game",
            "icon": "🎮",
            "color": COLORS['primary'],
            "completed": total_games >= 1,
            "progress": min(1.0, total_games / 1)
        },
        {
            "title": "Amateur",
            "description": "Win 5 games",
            "icon": "🥉",
            "color": COLORS['warning'],
            "completed": total_wins >= 5,
            "progress": min(1.0, total_wins / 5)
        },
        {
            "title": "Intermediate",
            "description": "Win 10 games",
            "icon": "🥈",
            "color": COLORS['warning'],
            "completed": total_wins >= 10,
            "progress": min(1.0, total_wins / 10)
        },
        {
            "title": "Expert",
            "description": "Win 25 games",
            "icon": "🥇",
            "color": COLORS['accent'],
            "completed": total_wins >= 25,
            "progress": min(1.0, total_wins / 25)
        },
        {
            "title": "Master",
            "description": "Win 50 games",
            "icon": "👑",
            "color": COLORS['primary'],
            "completed": total_wins >= 50,
            "progress": min(1.0, total_wins / 50)
        },
        {
            "title": "Dedicated Player",
            "description": "Play 100 games total",
            "icon": "🌟",
            "color": COLORS['accent'],
            "completed": total_games >= 100,
            "progress": min(1.0, total_games / 100)
        }
    ]
    
    # Create achievement cards
    for i, achievement in enumerate(achievements_list):
        card = ctk.CTkFrame(achievements_container, fg_color=COLORS['secondary_alt'], corner_radius=10)
        card.grid(row=i+1, column=0, padx=20, pady=5, sticky="ew")
        
        # Achievement icon
        icon_label = ctk.CTkLabel(card, 
                                text=achievement["icon"],
                                font=ctk.CTkFont(size=30),
                                text_color=achievement["color"])
        icon_label.grid(row=0, column=0, rowspan=2, padx=15, pady=15)
        
        # Achievement title with locked/unlocked status
        status = "✅" if achievement["completed"] else "🔒"
        title_text = f"{achievement['title']} {status}"
        title_label = ctk.CTkLabel(card,
                                 text=title_text,
                                 font=ctk.CTkFont(size=16, weight="bold"),
                                 text_color=COLORS['text'])
        title_label.grid(row=0, column=1, padx=10, pady=(15, 5), sticky="w")
        
        # Achievement description
        desc_label = ctk.CTkLabel(card,
                                text=achievement["description"],
                                font=ctk.CTkFont(size=13),
                                text_color=COLORS['text_secondary'])
        desc_label.grid(row=1, column=1, padx=10, pady=(0, 5), sticky="w")
        
        # Progress bar
        progress_bar = ctk.CTkProgressBar(card, height=10, corner_radius=5)
        progress_bar.grid(row=2, column=0, columnspan=2, padx=15, pady=(0, 15), sticky="ew")
        progress_bar.set(achievement["progress"])
        
        # Set progress bar color based on completion
        if achievement["completed"]:
            progress_bar.configure(progress_color=COLORS['accent'])
        else:
            progress_bar.configure(progress_color=COLORS['primary'])
        
        # Configure grid
        card.grid_columnconfigure(1, weight=1)

# Settings display
def update_settings_display():
    settings_frame = frames['settings']
    
    # Clear existing content
    for widget in settings_frame.winfo_children():
        widget.destroy()
    
    # Create header
    header = ctk.CTkFrame(settings_frame, fg_color=COLORS['secondary'], corner_radius=15)
    header.grid(row=0, column=0, sticky="ew", pady=(0, 20))
    
    title = ctk.CTkLabel(header, text="Game Settings",
                       font=ctk.CTkFont(family="Helvetica", size=24, weight="bold"),
                       text_color=COLORS['text'])
    title.grid(row=0, column=0, padx=20, pady=20, sticky="w")
    
    # Create settings container
    settings_container = ctk.CTkFrame(settings_frame, fg_color=COLORS['secondary'], corner_radius=15)
    settings_container.grid(row=1, column=0, sticky="ew", pady=(0, 20))
    settings_container.grid_columnconfigure(1, weight=1)
    
    # Player settings
    player_section = ctk.CTkLabel(settings_container, 
                                text="Player Settings",
                                font=ctk.CTkFont(family="Helvetica", size=18, weight="bold"),
                                text_color=COLORS['text'])
    player_section.grid(row=0, column=0, columnspan=2, padx=20, pady=(15, 15), sticky="w")
    
    # Player name
    name_label = ctk.CTkLabel(settings_container,
                            text="Player Name:",
                            font=ctk.CTkFont(size=14),
                            text_color=COLORS['text'])
    name_label.grid(row=1, column=0, padx=20, pady=10, sticky="w")
    
    name_entry = ctk.CTkEntry(settings_container, width=200)
    name_entry.grid(row=1, column=1, padx=20, pady=10, sticky="w")
    name_entry.insert(0, "Prithika")  # Default to current name
    
    # Game settings
    game_section = ctk.CTkLabel(settings_container, 
                              text="Game Settings",
                              font=ctk.CTkFont(family="Helvetica", size=18, weight="bold"),
                              text_color=COLORS['text'])
    game_section.grid(row=2, column=0, columnspan=2, padx=20, pady=(25, 15), sticky="w")
    
    # Sound effects toggle
    sound_label = ctk.CTkLabel(settings_container,
                             text="Sound Effects:",
                             font=ctk.CTkFont(size=14),
                             text_color=COLORS['text'])
    sound_label.grid(row=3, column=0, padx=20, pady=10, sticky="w")
    
    sound_switch = ctk.CTkSwitch(settings_container, text="", width=50)
    sound_switch.grid(row=3, column=1, padx=20, pady=10, sticky="w")
    sound_switch.select()  # Default to on
    
    # Difficulty level
    difficulty_label = ctk.CTkLabel(settings_container,
                                  text="Difficulty Level:",
                                  font=ctk.CTkFont(size=14),
                                  text_color=COLORS['text'])
    difficulty_label.grid(row=4, column=0, padx=20, pady=10, sticky="w")
    
    difficulty_options = ctk.CTkSegmentedButton(settings_container, 
                                             values=["Easy", "Medium", "Hard"])
    difficulty_options.grid(row=4, column=1, padx=20, pady=10, sticky="w")
    difficulty_options.set("Medium")  # Default to medium
    
    # Clear history
    clear_history_label = ctk.CTkLabel(settings_container,
                                     text="Game History:",
                                     font=ctk.CTkFont(size=14),
                                     text_color=COLORS['text'])
    clear_history_label.grid(row=5, column=0, padx=20, pady=10, sticky="w")
    
    clear_button = ctk.CTkButton(settings_container, 
                               text="Clear History",
                               fg_color=COLORS['danger'],
                               hover_color="#D55",
                               width=120)
    clear_button.grid(row=5, column=1, padx=20, pady=10, sticky="w")
    
    # Save button
    save_button = ctk.CTkButton(settings_container, 
                              text="Save Settings",
                              fg_color=COLORS['accent'],
                              hover_color="#2AB",
                              width=150)
    save_button.grid(row=6, column=0, columnspan=2, padx=20, pady=25)

# Help display
def update_help_display():
    help_frame = frames['help']
    
    # Clear existing content
    for widget in help_frame.winfo_children():
        widget.destroy()
    
    # Create header
    header = ctk.CTkFrame(help_frame, fg_color=COLORS['secondary'], corner_radius=15)
    header.grid(row=0, column=0, sticky="ew", pady=(0, 20))
    
    title = ctk.CTkLabel(header, text="Help & Instructions",
                       font=ctk.CTkFont(family="Helvetica", size=24, weight="bold"),
                       text_color=COLORS['text'])
    title.grid(row=0, column=0, padx=20, pady=20, sticky="w")
    
    # Create scrollable help content
    help_container = ctk.CTkScrollableFrame(help_frame, fg_color=COLORS['secondary'], corner_radius=15)
    help_container.grid(row=1, column=0, sticky="nsew", pady=(0, 20))
    help_container.grid_columnconfigure(0, weight=1)
    
    # Game rules section
    rules_title = ctk.CTkLabel(help_container, 
                             text="Game Rules",
                             font=ctk.CTkFont(family="Helvetica", size=18, weight="bold"),
                             text_color=COLORS['text'])
    rules_title.grid(row=0, column=0, padx=20, pady=(15, 10), sticky="w")
    
    rules_text = """
    Rock Paper Scissors is a simple game with the following rules:
    
    • Rock beats Scissors (Rock crushes Scissors)
    • Scissors beats Paper (Scissors cut Paper)
    • Paper beats Rock (Paper covers Rock)
    
    If both players choose the same option, it's a tie!
    """
    
    rules_label = ctk.CTkLabel(help_container,
                             text=rules_text,
                             font=ctk.CTkFont(size=14),
                             text_color=COLORS['text'],
                             justify="left",
                             wraplength=600)
    rules_label.grid(row=1, column=0, padx=20, pady=(0, 15), sticky="w")
    
    # How to play section
    how_to_title = ctk.CTkLabel(help_container, 
                              text="How to Play",
                              font=ctk.CTkFont(family="Helvetica", size=18, weight="bold"),
                              text_color=COLORS['text'])
    how_to_title.grid(row=2, column=0, padx=20, pady=(15, 10), sticky="w")
    
    how_to_text = """
    1. Go to the "Play Game" tab by clicking on the button in the sidebar
    2. Choose either Rock, Paper, or Scissors by clicking on the corresponding button
    3. The computer will randomly select its move
    4. The winner will be determined based on the game rules
    5. Your game history and statistics will be updated automatically
    """
    
    how_to_label = ctk.CTkLabel(help_container,
                              text=how_to_text,
                              font=ctk.CTkFont(size=14),
                              text_color=COLORS['text'],
                              justify="left",
                              wraplength=600)
    how_to_label.grid(row=3, column=0, padx=20, pady=(0, 15), sticky="w")
    
    # FAQ section
    faq_title = ctk.CTkLabel(help_container, 
                           text="Frequently Asked Questions",
                           font=ctk.CTkFont(family="Helvetica", size=18, weight="bold"),
                           text_color=COLORS['text'])
    faq_title.grid(row=4, column=0, padx=20, pady=(15, 10), sticky="w")
    
    # FAQ items
    faqs = [
        ("How are statistics calculated?", 
         "The game tracks your wins, losses, and ties. Win rate is calculated as (wins/total games) × 100%."),
        
        ("Can I reset my game history?", 
         "Yes, you can reset your game history in the Settings tab by clicking on 'Clear History'."),
        
        ("Is the computer's choice truly random?", 
         "Yes, the computer's choice is generated randomly with equal probability for Rock, Paper, and Scissors."),
        
        ("How do I earn achievements?", 
         "Achievements are earned automatically as you play more games and win more matches."),
        
        ("Can I change my player name?", 
         "Yes, you can change your name in the Settings tab.")
    ]
    
    for i, (question, answer) in enumerate(faqs):
        question_label = ctk.CTkLabel(help_container,
                                    text=question,
                                    font=ctk.CTkFont(size=14, weight="bold"),
                                    text_color=COLORS['primary'])
        question_label.grid(row=5+i*2, column=0, padx=20, pady=(10, 0), sticky="w")
        
        answer_label = ctk.CTkLabel(help_container,
                                  text=answer,
                                  font=ctk.CTkFont(size=14),
                                  text_color=COLORS['text'],
                                  wraplength=600,
                                  justify="left")
        answer_label.grid(row=6+i*2, column=0, padx=30, pady=(0, 5), sticky="w")
    
    # About section
    about_title = ctk.CTkLabel(help_container, 
                             text="About This Game",
                             font=ctk.CTkFont(family="Helvetica", size=18, weight="bold"),
                             text_color=COLORS['text'])
    about_title.grid(row=15, column=0, padx=20, pady=(25, 10), sticky="w")
    
    about_text = """
    Rock Paper Scissors Game v1.1.0
    
    This game was created for Prithika as a fun project to demonstrate Python GUI development
    using the CustomTkinter library.
    
    © 2025 RPS Game Development
    """
    
    about_label = ctk.CTkLabel(help_container,
                             text=about_text,
                             font=ctk.CTkFont(size=14),
                             text_color=COLORS['text_secondary'],
                             justify="left",
                             wraplength=600)
    about_label.grid(row=16, column=0, padx=20, pady=(0, 30), sticky="w")
