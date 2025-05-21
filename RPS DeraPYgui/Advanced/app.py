import dearpygui.dearpygui as dpg
import random
import time
import pandas as pd
import os
import sys

# Initialize DearPyGUI
dpg.create_context()

# Game constants
CHOICES = ["Rock", "Paper", "Scissors"]

# Excel file path
EXCEL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "rps_data.xlsx")

# Modern color theme
COLORS = {
    "background": [18, 18, 35],         # Dark blue-black
    "primary": [94, 84, 255],           # Purple
    "secondary": [64, 64, 99],          # Dark purple-blue
    "accent": [255, 122, 89],           # Coral
    "text": [240, 240, 250],            # Bright white with slight blue tint
    "win": [42, 209, 157],              # Teal green
    "lose": [255, 82, 99],              # Coral red
    "draw": [255, 203, 107],            # Amber yellow
    "panel": [28, 28, 50],              # Dark blue-purple
    "sidebar": [22, 22, 45],            # Darker blue-purple
    "sidebar_hover": [38, 38, 75],      # Lighter blue-purple
    "sidebar_active": [58, 48, 120],    # Highlight purple
    "card": [35, 35, 65],               # Card background
    "card_header": [45, 45, 85],        # Card header background
    "border": [55, 55, 90],             # Border color
    "chart1": [126, 87, 255],           # Chart color 1
    "chart2": [85, 208, 189],           # Chart color 2
    "chart3": [255, 168, 94]            # Chart color 3
}

# Game state
player_score = 0
computer_score = 0
game_history = []
round_count = 0
last_result = ""
total_rounds = 0
win_percentage = 0
draw_percentage = 0
current_view = "game"

# Layout constants for better proportions
LAYOUT = {
    "sidebar_width": 240,
    "padding": 10,
    "card_spacing": 15,
    "content_height": -1,
    "controls_width_ratio": 0.38,  # Controls take 38% of width
    "stats_width_ratio": 0.60,     # Stats take 60% of width
}

# Icons for choices and navigation - modern look
ICONS = {
    "Rock": "🎮",
    "Paper": "📝",
    "Scissors": "✂️",
    "Game": "🎲",
    "Stats": "📊",
    "History": "📜",
    "Settings": "⚙️",
    "Export": "📤",
    "Win": "🏆",
    "Lose": "❌",
    "Draw": "🔄"
}

# Function to save game data to Excel
def save_to_excel():
    # Create DataFrame from game history
    if not game_history:
        return
        
    history_data = []
    for i, h in enumerate(game_history):
        try:
            # More robust parsing of history entries
            round_info = h.split(": ", 1)  # Split only on first occurrence
            if len(round_info) < 2:
                continue  # Skip malformed entries
                
            timestamp_part = round_info[0]
            timestamp = timestamp_part.split("[")[1].split("]")[0] if "[" in timestamp_part and "]" in timestamp_part else "unknown"
            
            # Extract choices and result
            game_detail = round_info[1]
            
            # Handle different possible formats
            if " - " in game_detail:
                choices_part, result = game_detail.split(" - ", 1)
            else:
                choices_part, result = game_detail, "Unknown"
                
            # Extract player and computer choices
            if ", " in choices_part:
                player_part, computer_part = choices_part.split(", ", 1)
            else:
                player_part, computer_part = choices_part, "PC: Unknown"
                
            player_choice = player_part.replace("You: ", "") if "You: " in player_part else player_part
            computer_choice = computer_part.replace("PC: ", "") if "PC: " in computer_part else computer_part
            
            history_data.append({
                "Round": i+1,
                "Timestamp": timestamp,
                "Player Choice": player_choice,
                "Computer Choice": computer_choice,
                "Result": result
            })
        except Exception as e:
            print(f"Error parsing history entry {i}: {e}")
            # Add a fallback entry with available information
            history_data.append({
                "Round": i+1,
                "Timestamp": "error",
                "Player Choice": "error",
                "Computer Choice": "error",
                "Result": "Error parsing entry"
            })
    
    df_history = pd.DataFrame(history_data)
    
    # Create stats DataFrame
    stats_data = {
        "Statistic": ["Player Score", "Computer Score", "Total Rounds", "Win Rate", "Draw Rate"],
        "Value": [player_score, computer_score, total_rounds, 
                 f"{win_percentage:.1f}%" if total_rounds > 0 else "0%", 
                 f"{draw_percentage:.1f}%" if total_rounds > 0 else "0%"]
    }
    df_stats = pd.DataFrame(stats_data)
    
    # Save to Excel with multiple sheets
    try:
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            df_history.to_excel(writer, sheet_name='Game History', index=False)
            df_stats.to_excel(writer, sheet_name='Statistics', index=False)
        dpg.configure_item("status_text", default_value=f"Data saved to {EXCEL_FILE}", color=COLORS["win"])
    except Exception as e:
        dpg.configure_item("status_text", default_value=f"Error saving data: {str(e)}", color=COLORS["lose"])

# Function to load game data from Excel
def load_from_excel():
    global player_score, computer_score, game_history, round_count, total_rounds, win_percentage, draw_percentage
    
    if not os.path.exists(EXCEL_FILE):
        dpg.configure_item("status_text", default_value="No saved data found", color=COLORS["accent"])
        return False
    
    try:
        # Read history sheet
        df_history = pd.read_excel(EXCEL_FILE, sheet_name='Game History')
        
        # Convert history data back to our format
        game_history = []
        for _, row in df_history.iterrows():
            entry = f"Round {row['Round']} [{row['Timestamp']}]: You: {row['Player Choice']}, PC: {row['Computer Choice']} - {row['Result']}"
            game_history.append(entry)
        
        # Read stats sheet
        df_stats = pd.read_excel(EXCEL_FILE, sheet_name='Statistics')
        
        # Extract stats
        stats_dict = dict(zip(df_stats['Statistic'], df_stats['Value']))
        player_score = int(stats_dict.get('Player Score', 0))
        computer_score = int(stats_dict.get('Computer Score', 0))
        total_rounds = int(stats_dict.get('Total Rounds', 0))
        
        # Parse percentage values
        win_rate_str = stats_dict.get('Win Rate', '0%')
        win_percentage = float(win_rate_str.replace('%', ''))
        
        draw_rate_str = stats_dict.get('Draw Rate', '0%')
        draw_percentage = float(draw_rate_str.replace('%', ''))
        
        round_count = len(game_history)
        
        # Update UI
        update_displays()
        dpg.configure_item("history_list", items=game_history[::-1])
        dpg.configure_item("status_text", default_value="Data loaded successfully", color=COLORS["win"])
        
        return True
    except Exception as e:
        dpg.configure_item("status_text", default_value=f"Error loading data: {str(e)}", color=COLORS["lose"])
        return False

# Function to reset the game
def reset_game():
    global player_score, computer_score, game_history, round_count, last_result, total_rounds, win_percentage, draw_percentage
    
    # Reset game state variables
    player_score = 0
    computer_score = 0
    game_history = []
    round_count = 0
    last_result = ""
    total_rounds = 0
    win_percentage = 0
    draw_percentage = 0
    
    # Helper function to safely configure items
    def safe_configure_item(tag, **kwargs):
        try:
            if dpg.does_item_exist(tag):
                dpg.configure_item(tag, **kwargs)
        except Exception as e:
            print(f"Error configuring item {tag} in reset_game: {e}")
    
    # Update UI elements safely
    try:
        update_displays()
        safe_configure_item("history_list", items=[])
        safe_configure_item("detailed_history", items=[])
        safe_configure_item("result_text", default_value="Make your choice!")
        safe_configure_item("win_percentage", default_value="Win Rate: 0%")
        safe_configure_item("draw_percentage", default_value="Draw Rate: 0%")
        safe_configure_item("total_rounds_text", default_value="Total Rounds: 0")
        safe_configure_item("status_text", default_value="Game reset", color=COLORS["accent"])
        
        # Reset statistics view items if they exist
        safe_configure_item("stats_player_wins", default_value="0")
        safe_configure_item("stats_win_rate", default_value="0%")
        safe_configure_item("stats_computer_wins", default_value="0")
        safe_configure_item("stats_computer_win_rate", default_value="0%")
        safe_configure_item("stats_total_rounds", default_value="0")
        safe_configure_item("stats_draws", default_value="0")
        safe_configure_item("stats_draw_rate", default_value="0%")
        safe_configure_item("stats_fav_choice", default_value="N/A")
        
    except Exception as e:
        print(f"Error in reset_game: {e}")
        try:
            dpg.configure_item("status_text", default_value=f"Error resetting game: {str(e)}", color=COLORS["lose"])
        except:
            print("Could not update status text")

# Function to determine winner
def determine_winner(player, computer):
    if player == computer:
        return "Draw!"
    elif (player == "Rock" and computer == "Scissors") or \
         (player == "Paper" and computer == "Rock") or \
         (player == "Scissors" and computer == "Paper"):
        return "You win!"
    else:
        return "Computer wins!"

# Function to make a choice and play a round
def play_round(sender, app_data, user_data):
    global player_score, computer_score, game_history, round_count, last_result, total_rounds, win_percentage, draw_percentage
    
    # Get player choice and generate computer choice
    player_choice = user_data
    computer_choice = random.choice(CHOICES)
    
    # Determine the winner
    result = determine_winner(player_choice, computer_choice)
    last_result = f"You chose {player_choice}, Computer chose {computer_choice}."
    
    # Update scores
    if result == "You win!":
        player_score += 1
        result_color = COLORS["win"]
    elif result == "Computer wins!":
        computer_score += 1
        result_color = COLORS["lose"]
    else:
        result_color = COLORS["draw"]
    
    # Add to history (limit history size to prevent memory issues)
    MAX_HISTORY = 100
    round_count += 1
    total_rounds += 1
    timestamp = time.strftime("%H:%M:%S")
    history_entry = f"Round {round_count} [{timestamp}]: You: {player_choice}, PC: {computer_choice} - {result}"
    game_history.append(history_entry)
    
    # Trim history if it gets too large
    if len(game_history) > MAX_HISTORY:
        game_history = game_history[-MAX_HISTORY:]
    
    # Calculate statistics
    if total_rounds > 0:
        win_percentage = (player_score / total_rounds) * 100
        draw_count = total_rounds - player_score - computer_score
        draw_percentage = (draw_count / total_rounds) * 100
    
    # Update displays (use try-except to prevent crashes)
    try:
        update_displays(result, result_color)
        
        # Only update detailed history if history view is active
        if current_view == "history":
            dpg.configure_item("detailed_history", items=game_history[::-1])
    except Exception as e:
        print(f"Error updating displays: {e}")
        dpg.configure_item("status_text", default_value=f"Error: {str(e)}", color=COLORS["lose"])

# Function to update all displays
def update_displays(result="", result_color=COLORS["text"]):
    try:
        # Helper function to safely configure items
        def safe_configure_item(tag, **kwargs):
            try:
                if dpg.does_item_exist(tag):
                    dpg.configure_item(tag, **kwargs)
            except Exception as e:
                print(f"Error configuring item {tag}: {e}")
        
        # Update result
        if not result:
            safe_configure_item("result_text", default_value=last_result)
        else:
            safe_configure_item("result_text", default_value=last_result)
            safe_configure_item("result_outcome", default_value=result, color=result_color)
        
        # Update score
        player_color = COLORS["win"] if player_score > computer_score else COLORS["text"]
        computer_color = COLORS["lose"] if player_score > computer_score else COLORS["text"]
        if player_score == computer_score:
            player_color = computer_color = COLORS["draw"]
        
        safe_configure_item("player_score", default_value=str(player_score), color=player_color)
        safe_configure_item("computer_score", default_value=str(computer_score), color=computer_color)
        
        # Update statistics
        safe_configure_item("win_percentage", default_value=f"Win Rate: {win_percentage:.1f}%")
        safe_configure_item("draw_percentage", default_value=f"Draw Rate: {draw_percentage:.1f}%")
        safe_configure_item("total_rounds_text", default_value=f"Total Rounds: {total_rounds}")
        
        # Update progress bars for win and draw rates
        if total_rounds > 0:
            dpg.configure_item("win_rate_bar", default_value=win_percentage/100, overlay=f"{win_percentage:.1f}%")
            dpg.configure_item("draw_rate_bar", default_value=draw_percentage/100, overlay=f"{draw_percentage:.1f}%")
        
        # Only update history in game view (detailed history updated separately)
        if game_history:
            safe_configure_item("history_list", items=game_history[-8:][::-1])  # Show only last 8 entries for better performance
    except Exception as e:
        print(f"Error in update_displays: {e}")

# Function to switch views
def switch_view(sender, app_data, user_data):
    global current_view
    current_view = user_data
    
    try:
        # Hide all views
        dpg.configure_item("game_view", show=False)
        dpg.configure_item("stats_view", show=False)
        dpg.configure_item("history_view", show=False)
        dpg.configure_item("settings_view", show=False)
        
        # Show selected view
        dpg.configure_item(f"{user_data}_view", show=True)
        
        # Update sidebar buttons
        for view in ["game", "stats", "history", "settings"]:
            if view == user_data:
                dpg.bind_item_theme(f"{view}_button", active_button_theme)
            else:
                dpg.bind_item_theme(f"{view}_button", button_theme)
        
        # Update view-specific data
        if user_data == "history":
            dpg.configure_item("detailed_history", items=game_history[::-1])
        elif user_data == "stats" and total_rounds > 0:
            update_statistics_view()
    except Exception as e:
        print(f"Error switching views: {e}")
        dpg.configure_item("status_text", default_value=f"Error: {str(e)}", color=COLORS["lose"])

# New function to update stats view separately (for better performance)
def update_statistics_view():
    if total_rounds == 0:
        return
        
    try:
        dpg.configure_item("stats_player_wins", default_value=str(player_score))
        dpg.configure_item("stats_win_rate", default_value=f"{win_percentage:.1f}%")
        
        dpg.configure_item("stats_computer_wins", default_value=str(computer_score))
        computer_win_rate = (computer_score / total_rounds) * 100
        dpg.configure_item("stats_computer_win_rate", default_value=f"{computer_win_rate:.1f}%")
        
        dpg.configure_item("stats_total_rounds", default_value=str(total_rounds))
        draws = total_rounds - player_score - computer_score
        dpg.configure_item("stats_draws", default_value=str(draws))
        dpg.configure_item("stats_draw_rate", default_value=f"{draw_percentage:.1f}%")
        
        # Calculate favorite choice if there's history
        if game_history:
            choices_count = {"Rock": 0, "Paper": 0, "Scissors": 0}
            # Only process the last 50 games for performance
            for entry in game_history[-50:]:
                for choice in CHOICES:
                    if f"You: {choice}" in entry:
                        choices_count[choice] += 1
            
            favorite = max(choices_count, key=choices_count.get)
            dpg.configure_item("stats_fav_choice", default_value=f"{favorite} ({choices_count[favorite]} times)")
    except Exception as e:
        print(f"Error updating statistics: {e}")

# Create modern theme with rounded corners and better spacing
with dpg.theme() as global_theme:
    with dpg.theme_component(dpg.mvAll):
        dpg.add_theme_color(dpg.mvThemeCol_WindowBg, COLORS["background"])
        dpg.add_theme_color(dpg.mvThemeCol_TitleBgActive, COLORS["primary"])
        dpg.add_theme_color(dpg.mvThemeCol_Button, COLORS["secondary"])
        dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, COLORS["primary"])
        dpg.add_theme_color(dpg.mvThemeCol_ButtonActive, COLORS["accent"])
        dpg.add_theme_color(dpg.mvThemeCol_Text, COLORS["text"])
        dpg.add_theme_color(dpg.mvThemeCol_FrameBg, COLORS["panel"])
        dpg.add_theme_color(dpg.mvThemeCol_Border, COLORS["border"])
        dpg.add_theme_color(dpg.mvThemeCol_ScrollbarBg, COLORS["panel"])
        dpg.add_theme_color(dpg.mvThemeCol_ScrollbarGrab, COLORS["secondary"])
        dpg.add_theme_color(dpg.mvThemeCol_ScrollbarGrabHovered, COLORS["primary"])
        dpg.add_theme_color(dpg.mvThemeCol_ScrollbarGrabActive, COLORS["accent"])
        dpg.add_theme_color(dpg.mvThemeCol_Header, COLORS["card_header"])
        dpg.add_theme_color(dpg.mvThemeCol_HeaderHovered, COLORS["primary"])
        dpg.add_theme_color(dpg.mvThemeCol_TableHeaderBg, COLORS["card_header"])
        dpg.add_theme_style(dpg.mvStyleVar_FrameRounding, 10)
        dpg.add_theme_style(dpg.mvStyleVar_WindowRounding, 12)
        dpg.add_theme_style(dpg.mvStyleVar_ChildRounding, 12)
        dpg.add_theme_style(dpg.mvStyleVar_FramePadding, 10, 8)
        dpg.add_theme_style(dpg.mvStyleVar_ItemSpacing, 12, 8)
        dpg.add_theme_style(dpg.mvStyleVar_ItemInnerSpacing, 8, 6)
        dpg.add_theme_style(dpg.mvStyleVar_WindowTitleAlign, 0.5, 0.5)
        dpg.add_theme_style(dpg.mvStyleVar_ScrollbarSize, 15)
        dpg.add_theme_style(dpg.mvStyleVar_ScrollbarRounding, 9)
        dpg.add_theme_style(dpg.mvStyleVar_GrabMinSize, 20)
        dpg.add_theme_style(dpg.mvStyleVar_GrabRounding, 5)

# Create modern button themes with gradients
with dpg.theme() as button_theme:
    with dpg.theme_component(dpg.mvButton):
        dpg.add_theme_color(dpg.mvThemeCol_Button, COLORS["sidebar"])
        dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, COLORS["sidebar_hover"])
        dpg.add_theme_color(dpg.mvThemeCol_ButtonActive, COLORS["sidebar_active"])
        dpg.add_theme_style(dpg.mvStyleVar_FrameRounding, 8)
        dpg.add_theme_style(dpg.mvStyleVar_FramePadding, 8, 6)

with dpg.theme() as active_button_theme:
    with dpg.theme_component(dpg.mvButton):
        dpg.add_theme_color(dpg.mvThemeCol_Button, COLORS["primary"])
        dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, COLORS["primary"])
        dpg.add_theme_color(dpg.mvThemeCol_ButtonActive, COLORS["accent"])
        dpg.add_theme_style(dpg.mvStyleVar_FrameRounding, 8)
        dpg.add_theme_style(dpg.mvStyleVar_FramePadding, 8, 6)

# Card theme for dashboard panels
with dpg.theme() as card_theme:
    with dpg.theme_component(dpg.mvChildWindow):
        dpg.add_theme_color(dpg.mvThemeCol_ChildBg, COLORS["card"])
        dpg.add_theme_color(dpg.mvThemeCol_Border, COLORS["border"])
        dpg.add_theme_style(dpg.mvStyleVar_ChildRounding, 12)
        dpg.add_theme_style(dpg.mvStyleVar_ChildBorderSize, 1)
        dpg.add_theme_style(dpg.mvStyleVar_WindowPadding, 15, 15)

# Action button theme
with dpg.theme() as action_button_theme:
    with dpg.theme_component(dpg.mvButton):
        dpg.add_theme_color(dpg.mvThemeCol_Button, COLORS["primary"])
        dpg.add_theme_color(dpg.mvThemeCol_ButtonHovered, [c + 30 for c in COLORS["primary"]])
        dpg.add_theme_color(dpg.mvThemeCol_ButtonActive, COLORS["accent"])
        dpg.add_theme_style(dpg.mvStyleVar_FrameRounding, 20)
        dpg.add_theme_style(dpg.mvStyleVar_FramePadding, 10, 8)

# Create viewport with modern title
dpg.create_viewport(title="Rock Paper Scissors | Modern Dashboard", width=1200, height=800)
dpg.bind_theme(global_theme)
dpg.configure_viewport(0, maximized=True)

# Custom fonts for modern look
with dpg.font_registry():
    # Load default font
    default_font = dpg.add_font("segoeui.ttf", 16) if os.path.exists("segoeui.ttf") else None
    
    # Create larger font for score display
    large_font = dpg.add_font("segoeui.ttf", 24) if os.path.exists("segoeui.ttf") else None
    
    # Bind default font if available
    if default_font:
        dpg.bind_font(default_font)

# Main window - now uses full space with optimal alignment
with dpg.window(label="Rock Paper Scissors Dashboard", tag="primary_window", no_resize=True, no_close=True, no_collapse=True, no_move=True):
    # App title - centered with better spacing
    with dpg.group():
        dpg.add_spacer(height=5)
        with dpg.group(horizontal=True):
            title_text = dpg.add_text("ROCK PAPER SCISSORS", color=COLORS["primary"])
            dpg.add_text("DASHBOARD", color=COLORS["accent"])
            # Center title using a dummy spacer after calculating width
            title_width = 250  # Approximate width of the title
            screen_width = dpg.get_viewport_width()
            if screen_width > title_width:
                dpg.add_dummy(width=(screen_width - title_width) // 2)
        dpg.add_spacer(height=10)
    
    # Main layout with improved proportions (sidebar + content area)
    with dpg.group(horizontal=True):
        # Sidebar with fixed width
        with dpg.child_window(width=LAYOUT["sidebar_width"], tag="sidebar", border=False, height=LAYOUT["content_height"]):
            dpg.add_spacer(height=10)
            
            # User profile section (decorative)
            with dpg.group(horizontal=True):
                dpg.add_text("👤", color=COLORS["accent"])
                dpg.add_text("PLAYER", color=COLORS["text"])
            
            dpg.add_separator()
            dpg.add_spacer(height=15)
            
            # Navigation buttons - full width with consistent spacing
            dpg.add_text("DASHBOARD", color=COLORS["accent"], indent=5)
            dpg.add_spacer(height=8)
            
            with dpg.group():
                dpg.add_button(label=f"{ICONS['Game']} Game", callback=switch_view, user_data="game", width=-1, height=45, tag="game_button")
                dpg.bind_item_theme("game_button", active_button_theme)
                
                dpg.add_button(label=f"{ICONS['Stats']} Statistics", callback=switch_view, user_data="stats", width=-1, height=45, tag="stats_button")
                dpg.bind_item_theme("stats_button", button_theme)
                
                dpg.add_button(label=f"{ICONS['History']} History", callback=switch_view, user_data="history", width=-1, height=45, tag="history_button")
                dpg.bind_item_theme("history_button", button_theme)
                
                dpg.add_button(label=f"{ICONS['Settings']} Settings", callback=switch_view, user_data="settings", width=-1, height=45, tag="settings_button")
                dpg.bind_item_theme("settings_button", button_theme)
            
            dpg.add_spacer(height=20)
            
            # Data options section - full width
            dpg.add_text("DATA OPTIONS", color=COLORS["accent"], indent=5)
            dpg.add_spacer(height=8)
            
            dpg.add_button(label=f"{ICONS['Export']} Save to Excel", callback=save_to_excel, width=-1, height=40)
            dpg.add_button(label="📥 Load from Excel", callback=load_from_excel, width=-1, height=40)
            
            dpg.add_spacer(height=20)
            
            # Actions section - full width with increased height
            dpg.add_text("ACTIONS", color=COLORS["accent"], indent=5)
            dpg.add_spacer(height=8)
            
            dpg.add_button(label="🔄 Reset Game", callback=reset_game, width=-1, height=40)
            dpg.add_button(label="❌ Exit", callback=lambda: sys.exit(0), width=-1, height=40)
            
            # Fill remaining space before status
            dpg.add_spacer(height=20)
            dpg.add_separator()
            
            # Status text at the bottom of sidebar - better positioned
            with dpg.group(horizontal=True):
                dpg.add_text("⚡", color=COLORS["accent"])
                dpg.add_text("Status:", color=COLORS["accent"])
            dpg.add_text("Ready to play", tag="status_text", color=COLORS["text"], wrap=LAYOUT["sidebar_width"]-20, indent=10)
        
        # Content Area - takes all remaining space
        with dpg.child_window(tag="content_area", border=False, width=-1):
            # Game View - optimized layout that fills available space
            with dpg.child_window(tag="game_view", show=True, border=False):
                # Game title with session info
                with dpg.group(horizontal=True):
                    dpg.add_text(f"{ICONS['Game']} GAMEPLAY", color=COLORS["primary"])
                    dpg.add_spacer(width=20)
                    dpg.add_text(f"SESSION: {time.strftime('%d %b %Y')}", color=COLORS["accent"])
                
                dpg.add_spacer(height=LAYOUT["padding"])
                
                # Dynamic width calculation for better space usage
                viewport_width = dpg.get_viewport_width()
                content_width = viewport_width - LAYOUT["sidebar_width"] - 30  # Account for padding
                
                # Game control width based on ratio
                controls_width = int(content_width * LAYOUT["controls_width_ratio"])
                stats_width = int(content_width * LAYOUT["stats_width_ratio"])
                
                # Main game content - better proportioned horizontal layout
                with dpg.group(horizontal=True):
                    # Game controls card - proportional width
                    with dpg.child_window(width=controls_width, height=400, tag="game_controls_card"):
                        dpg.bind_item_theme("game_controls_card", card_theme)
                        
                        with dpg.group():
                            with dpg.group(horizontal=True):
                                dpg.add_text("🎮", color=COLORS["accent"])
                                dpg.add_text("MAKE YOUR CHOICE", color=COLORS["primary"])
                            
                            dpg.add_separator()
                            dpg.add_spacer(height=15)
                            
                            # Choice buttons with consistent spacing
                            for choice in CHOICES:
                                dpg.add_button(label=f"{ICONS[choice]} {choice}", callback=play_round, 
                                              user_data=choice, width=-1, height=80)
                                dpg.add_spacer(height=12)
                            
                            # Last result display with full width
                            with dpg.group():
                                dpg.add_text("", tag="result_text", color=COLORS["text"], wrap=controls_width-30)
                                dpg.add_text("", tag="result_outcome", color=COLORS["text"], wrap=controls_width-30)
                    
                    dpg.add_spacer(width=LAYOUT["card_spacing"])
                    
                    # Right column with score and stats - takes remaining space
                    with dpg.group(width=stats_width):
                        # Score card - full available width
                        with dpg.child_window(width=-1, height=180, tag="score_card"):
                            dpg.bind_item_theme("score_card", card_theme)
                            
                            with dpg.group():
                                with dpg.group(horizontal=True):
                                    dpg.add_text("🏆", color=COLORS["accent"])
                                    dpg.add_text("SCOREBOARD", color=COLORS["primary"])
                                
                                dpg.add_separator()
                                dpg.add_spacer(height=15)
                                
                                # Modern score display with equal proportions
                                score_item_width = (stats_width - 60) // 3
                                
                                with dpg.group(horizontal=True):
                                    # Player score with card
                                    with dpg.child_window(width=score_item_width, height=90, tag="player_score_card"):
                                        player_label = dpg.add_text("YOU", color=COLORS["text"])
                                        dpg.set_item_pos(player_label, [(score_item_width-30)//2, 10])
                                        
                                        player_score_text = dpg.add_text("0", tag="player_score", color=COLORS["win"])
                                        dpg.set_item_pos(player_score_text, [(score_item_width-10)//2, 40])
                                        if large_font:
                                            dpg.bind_item_font(player_score_text, large_font)
                                    
                                    # VS separator - centered
                                    with dpg.child_window(width=score_item_width, height=90):
                                        vs_text = dpg.add_text("VS", color=COLORS["accent"])
                                        dpg.set_item_pos(vs_text, [(score_item_width-20)//2, 40])
                                    
                                    # Computer score with card
                                    with dpg.child_window(width=score_item_width, height=90, tag="computer_score_card"):
                                        cpu_label = dpg.add_text("CPU", color=COLORS["text"])
                                        dpg.set_item_pos(cpu_label, [(score_item_width-30)//2, 10])
                                        
                                        computer_score_text = dpg.add_text("0", tag="computer_score", color=COLORS["lose"])
                                        dpg.set_item_pos(computer_score_text, [(score_item_width-10)//2, 40])
                                        if large_font:
                                            dpg.bind_item_font(computer_score_text, large_font)
                        
                        dpg.add_spacer(height=LAYOUT["card_spacing"])
                        
                        # Stats Summary Card - full available width
                        with dpg.child_window(width=-1, height=200, tag="quick_stats_card"):
                            dpg.bind_item_theme("quick_stats_card", card_theme)
                            
                            with dpg.group():
                                with dpg.group(horizontal=True):
                                    dpg.add_text("📊", color=COLORS["accent"])
                                    dpg.add_text("QUICK STATS", color=COLORS["primary"])
                                
                                dpg.add_separator()
                                dpg.add_spacer(height=15)
                                
                                # Stats in a more visual format with full-width progress bars
                                dpg.add_text("Win Rate:", color=COLORS["win"])
                                dpg.add_progress_bar(default_value=0, width=-1, height=25, tag="win_rate_bar", overlay="0%")
                                
                                dpg.add_spacer(height=12)
                                
                                dpg.add_text("Draw Rate:", color=COLORS["draw"])
                                dpg.add_progress_bar(default_value=0, width=-1, height=25, tag="draw_rate_bar", overlay="0%")
                                
                                dpg.add_spacer(height=12)
                                
                                with dpg.group(horizontal=True):
                                    dpg.add_text("Total Rounds:", color=COLORS["text"])
                                    dpg.add_spacer(width=10)
                                    dpg.add_text("0", tag="total_rounds_text", color=COLORS["accent"])
                
                dpg.add_spacer(height=LAYOUT["card_spacing"])
                
                # Recent History Card - full width
                with dpg.child_window(width=-1, height=200, tag="history_card"):
                    dpg.bind_item_theme("history_card", card_theme)
                    
                    with dpg.group():
                        with dpg.group(horizontal=True):
                            dpg.add_text("📜", color=COLORS["accent"])
                            dpg.add_text("RECENT MATCHES", color=COLORS["primary"])
                        
                        dpg.add_separator()
                        dpg.add_spacer(height=10)
                        
                        # History display with increased height for more entries
                        dpg.add_listbox(items=[], tag="history_list", width=-1, num_items=7)
            
            # Statistics View
            with dpg.child_window(tag="stats_view", show=False, border=False):
                dpg.add_text("GAME STATISTICS", color=COLORS["accent"])
                dpg.add_separator()
                
                with dpg.group(horizontal=True):
                    # Player stats
                    with dpg.child_window(width=350, height=300, label="Player Stats"):
                        dpg.add_text("Player Performance", color=COLORS["accent"])
                        dpg.add_separator()
                        
                        with dpg.group():
                            with dpg.group(horizontal=True):
                                dpg.add_text("Wins:", color=COLORS["win"])
                                dpg.add_spacer(width=10)
                                dpg.add_text("0", tag="stats_player_wins", color=COLORS["win"])
                            
                            with dpg.group(horizontal=True):
                                dpg.add_text("Win Rate:", color=COLORS["win"])
                                dpg.add_spacer(width=10)
                                dpg.add_text("0%", tag="stats_win_rate", color=COLORS["win"])
                            

                            with dpg.group(horizontal=True):
                                dpg.add_text("Favorite Choice:", color=COLORS["text"])
                                dpg.add_spacer(width=10)
                                dpg.add_text("N/A", tag="stats_fav_choice", color=COLORS["text"])
                    
                    # Computer stats
                    with dpg.child_window(width=350, height=300, label="Computer Stats"):
                        dpg.add_text("Computer Performance", color=COLORS["accent"])
                        dpg.add_separator()
                        
                        with dpg.group():
                            with dpg.group(horizontal=True):
                                dpg.add_text("Wins:", color=COLORS["lose"])
                                dpg.add_spacer(width=10)
                                dpg.add_text("0", tag="stats_computer_wins", color=COLORS["lose"])
                            
                            with dpg.group(horizontal=True):
                                dpg.add_text("Win Rate:", color=COLORS["lose"])
                                dpg.add_spacer(width=10)
                                dpg.add_text("0%", tag="stats_computer_win_rate", color=COLORS["lose"])
                
                # Game summary
                with dpg.child_window(width=-1, height=200, label="Game Summary"):
                    dpg.add_text("Overall Game Statistics", color=COLORS["accent"])
                    dpg.add_separator()
                    
                    with dpg.table(header_row=True):
                        dpg.add_table_column(label="Statistic")
                        dpg.add_table_column(label="Value")
                        
                        with dpg.table_row():
                            dpg.add_text("Total Rounds")
                            dpg.add_text("0", tag="stats_total_rounds")
                        
                        with dpg.table_row():
                            dpg.add_text("Draws")
                            dpg.add_text("0", tag="stats_draws")
                        
                        with dpg.table_row():
                            dpg.add_text("Draw Rate")
                            dpg.add_text("0%", tag="stats_draw_rate")
            
            # History View
            with dpg.child_window(tag="history_view", show=False, border=False):
                dpg.add_text("GAME HISTORY", color=COLORS["accent"])
                dpg.add_separator()
                
                # Full history display with filtering options
                dpg.add_text("Complete Match History:", color=COLORS["text"])
                dpg.add_listbox(items=[], tag="detailed_history", width=-1, num_items=15)
                
                # Export options
                with dpg.group(horizontal=True):
                    dpg.add_button(label="Export History to Excel", callback=save_to_excel, width=200, height=30)
                    dpg.add_spacer(width=10)
                    dpg.add_button(label="Clear History", callback=reset_game, width=150, height=30)
            
            # Settings View
            with dpg.child_window(tag="settings_view", show=False, border=False):
                dpg.add_text("SETTINGS", color=COLORS["accent"])
                dpg.add_separator()
                
                dpg.add_text("Game Settings", color=COLORS["text"])
                
                # Placeholder for future settings
                dpg.add_text("Settings will be available in future updates.")
                
                dpg.add_separator()
                
                # About section
                dpg.add_text("About this Application", color=COLORS["accent"])
                dpg.add_text("Rock Paper Scissors Dashboard")
                dpg.add_text("A simple game with statistics tracking and Excel export.")

# Setup and start
dpg.setup_dearpygui()
dpg.show_viewport()
dpg.set_primary_window("primary_window", True)

# Main application loop
last_stats_update = time.time()
update_interval = 0.5  # Update stats every 0.5 seconds at most

while dpg.is_dearpygui_running():
    try:
        # Only update stats periodically instead of every frame
        current_time = time.time()
        if current_view == "stats" and total_rounds > 0 and (current_time - last_stats_update) > update_interval:
            update_statistics_view()
            last_stats_update = current_time
            
        dpg.render_dearpygui_frame()
    except Exception as e:
        print(f"Error in main loop: {e}")
        time.sleep(0.1)  # Give the system some time to recover if there's an error
