import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np

def create_charts(parent_frame):
    # Create a figure with a dark background
    fig = plt.figure(figsize=(12, 4), facecolor='#1E1E2E')
    
    # Add subplots
    moves_ax = plt.subplot(121)
    trend_ax = plt.subplot(122)
    
    # Color scheme for charts - matching the app's dark theme
    chart_colors = {
        'bg': '#1E1E2E',
        'rock': '#F08080',   # Light coral
        'paper': '#8C9EFF',  # Light indigo
        'scissors': '#98FB98', # Light green
        'trend': '#8B5CF6',  # Purple
        'text': '#FFFFFF',   # White
        'grid': '#3F3F5F'    # Dark blue-gray
    }
    
    def update_charts(game_data):
        # Clear previous plots
        moves_ax.clear()
        trend_ax.clear()
        
        # Update moves distribution chart with enhanced styling
        moves = game_data['moves']
        
        # Skip empty data
        if sum(moves.values()) == 0:
            moves_ax.text(0.5, 0.5, 'No data available', 
                         horizontalalignment='center',
                         verticalalignment='center',
                         transform=moves_ax.transAxes,
                         color=chart_colors['text'], fontsize=12)
        else:
            wedges, texts, autotexts = moves_ax.pie(
                moves.values(), 
                labels=None, 
                autopct='%1.1f%%',
                colors=[chart_colors['rock'], chart_colors['paper'], chart_colors['scissors']],
                wedgeprops={'edgecolor': chart_colors['bg'], 'linewidth': 2, 'antialiased': True}
            )
            
            # Customize the appearance of percentage text
            for autotext in autotexts:
                autotext.set_fontsize(10)
                autotext.set_weight('bold')
                autotext.set_color('#1E1E2E')
            
            # Add a legend instead of labels
            moves_ax.legend(
                wedges, 
                moves.keys(),
                title="Move Types",
                loc="center left",
                bbox_to_anchor=(0.9, 0, 0.5, 1),
                fontsize=9,
                frameon=False,
                title_fontsize=10
            )
        
        moves_ax.set_title('Move Distribution', color=chart_colors['text'], fontsize=12, pad=20)
        
        # Update win rate trend with enhanced styling
        trend = game_data['trend']
        
        if len(trend) <= 1:
            trend_ax.text(0.5, 0.5, 'Not enough games played', 
                         horizontalalignment='center',
                         verticalalignment='center',
                         transform=trend_ax.transAxes,
                         color=chart_colors['text'], fontsize=12)
        else:
            # Create x values for smooth plotting
            x = np.arange(len(trend))
            
            # Create gradient line
            trend_ax.plot(x, trend, color=chart_colors['trend'], linewidth=2.5, alpha=0.9)
            
            # Add subtle grid
            trend_ax.grid(True, linestyle='--', alpha=0.2, color=chart_colors['grid'])
            
            # Add subtle gradient background to highlight the trend
            trend_ax.fill_between(x, trend, alpha=0.2, color=chart_colors['trend'])
            
            # Add labels for the latest win rate
            if len(trend) > 1:
                latest_rate = trend[-1]
                trend_ax.annotate(f'{latest_rate:.1f}%', 
                                 xy=(len(trend)-1, latest_rate),
                                 xytext=(10, 0),
                                 textcoords="offset points",
                                 color=chart_colors['text'],
                                 fontweight='bold')
            
            trend_ax.set_ylim(0, 105)  # Leave a bit of space at the top
            
            # Remove x ticks but keep y ticks
            trend_ax.set_xticks([])
            trend_ax.tick_params(axis='y', colors=chart_colors['text'])
        
        trend_ax.set_title('Win Rate Trend', color=chart_colors['text'], fontsize=12, pad=20)
        
        # Style adjustments
        for ax in [moves_ax, trend_ax]:
            ax.set_facecolor(chart_colors['bg'])
            
        fig.tight_layout()
        canvas.draw()
    
    # Create canvas
    canvas = FigureCanvasTkAgg(fig, master=parent_frame)
    widget = canvas.get_tk_widget()
    
    # Store update function as attribute
    widget.update_charts = update_charts
    
    return widget
