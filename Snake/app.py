import dearpygui.dearpygui as dpg
import random
import time

# Initialize DearPyGUI
dpg.create_context()

# Game constants
GRID_SIZE = 20
CELL_SIZE = 20
GRID_WIDTH = 30
GRID_HEIGHT = 20
GAME_SPEED = 0.1  # seconds per update

# Game state
game_running = False
last_update_time = 0
snake = [(5, 5), (5, 6), (5, 7)]  # List of (x, y) tuples, head is first element
direction = (1, 0)  # (dx, dy)
food_position = (10, 10)
score = 0
game_over = False

# Function to reset the game
def reset_game():
    global snake, direction, food_position, score, game_over, game_running
    snake = [(5, 5), (5, 6), (5, 7)]
    direction = (1, 0)
    spawn_food()
    score = 0
    game_over = False
    game_running = False
    dpg.set_value("score_text", f"Score: {score}")
    dpg.configure_item("game_over_text", show=False)

# Function to spawn food at random position
def spawn_food():
    global food_position
    while True:
        x = random.randint(0, GRID_WIDTH - 1)
        y = random.randint(0, GRID_HEIGHT - 1)
        if (x, y) not in snake:
            food_position = (x, y)
            break

# Function to handle key presses
def handle_key_press(sender, app_data):
    global direction, game_running
    
    if game_over:
        if app_data == dpg.mvKey_R:
            reset_game()
        return
    
    key = app_data
    
    # Start game on any key press if not already running
    if not game_running:
        game_running = True
        return
    
    # Change direction based on key press
    if key == dpg.mvKey_Up and direction != (0, 1):
        direction = (0, -1)
    elif key == dpg.mvKey_Down and direction != (0, -1):
        direction = (0, 1)
    elif key == dpg.mvKey_Left and direction != (1, 0):
        direction = (-1, 0)
    elif key == dpg.mvKey_Right and direction != (-1, 0):
        direction = (1, 0)

# Function to update game state
def update_game():
    global snake, food_position, score, game_over, last_update_time
    
    if not game_running or game_over:
        return
    
    current_time = time.time()
    if current_time - last_update_time < GAME_SPEED:
        return
    
    last_update_time = current_time
    
    # Move snake
    head_x, head_y = snake[0]
    dx, dy = direction
    new_head = (head_x + dx, head_y + dy)
    
    # Check for collisions with walls
    if (new_head[0] < 0 or new_head[0] >= GRID_WIDTH or
        new_head[1] < 0 or new_head[1] >= GRID_HEIGHT):
        game_over = True
        dpg.configure_item("game_over_text", show=True)
        return
    
    # Check for collisions with self
    if new_head in snake:
        game_over = True
        dpg.configure_item("game_over_text", show=True)
        return
    
    # Move the snake forward
    snake.insert(0, new_head)
    
    # Check if snake ate food
    if new_head == food_position:
        score += 1
        dpg.set_value("score_text", f"Score: {score}")
        spawn_food()
    else:
        snake.pop()  # Remove tail if no food was eaten

# Function to render the game
def render_game():
    dpg.delete_item("game_canvas", children_only=True)
    
    # Draw grid
    for x in range(GRID_WIDTH):
        for y in range(GRID_HEIGHT):
            dpg.draw_rectangle(
                (x * CELL_SIZE, y * CELL_SIZE),
                ((x + 1) * CELL_SIZE, (y + 1) * CELL_SIZE),
                fill=(40, 40, 40),
                color=(50, 50, 50),
                parent="game_canvas"
            )
    
    # Draw snake
    for i, (x, y) in enumerate(snake):
        color = (0, 200, 0) if i == 0 else (0, 255, 0)  # Head is darker green
        dpg.draw_rectangle(
            (x * CELL_SIZE, y * CELL_SIZE),
            ((x + 1) * CELL_SIZE, (y + 1) * CELL_SIZE),
            fill=color,
            color=color,
            parent="game_canvas"
        )
    
    # Draw food
    food_x, food_y = food_position
    dpg.draw_rectangle(
        (food_x * CELL_SIZE, food_y * CELL_SIZE),
        ((food_x + 1) * CELL_SIZE, (food_y + 1) * CELL_SIZE),
        fill=(255, 0, 0),
        color=(255, 0, 0),
        parent="game_canvas"
    )

# Function to start a new game
def start_game():
    reset_game()
    global game_running
    game_running = True

# Main game loop
def game_loop():
    update_game()
    render_game()

# Create viewport and window
dpg.create_viewport(title="Snake Game", width=GRID_WIDTH * CELL_SIZE + 40, height=GRID_HEIGHT * CELL_SIZE + 100)

with dpg.window(label="Snake Game", tag="primary_window", no_resize=True):
    dpg.add_text("Welcome to Snake Game!", color=[0, 255, 0])
    dpg.add_text("Use arrow keys to control the snake.", color=[200, 200, 200])
    dpg.add_text("Press any key to start!", color=[200, 200, 200])
    dpg.add_text("Score: 0", tag="score_text")
    dpg.add_text("GAME OVER! Press 'R' to restart", tag="game_over_text", color=[255, 0, 0], show=False)
    
    dpg.add_button(label="New Game", callback=start_game)
    
    with dpg.drawlist(width=GRID_WIDTH * CELL_SIZE, height=GRID_HEIGHT * CELL_SIZE, tag="game_canvas"):
        pass

# Register key handler
with dpg.handler_registry():
    dpg.add_key_press_handler(callback=handle_key_press)

# Setup and start
dpg.setup_dearpygui()
dpg.show_viewport()
dpg.set_primary_window("primary_window", True)

# Initialize game state
spawn_food()

# Main application loop
while dpg.is_dearpygui_running():
    game_loop()
    dpg.render_dearpygui_frame()

dpg.destroy_context()
