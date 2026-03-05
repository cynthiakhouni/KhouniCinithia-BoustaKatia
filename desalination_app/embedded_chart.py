"""
Embedded Interactive Chart Component using matplotlib
Provides tooltips, zoom, pan, and legend toggle within tkinter app.
"""

import tkinter as tk
from tkinter import ttk
from typing import List, Dict, Tuple, Optional
from datetime import datetime, timedelta

# Try to import matplotlib components
try:
    import matplotlib
    matplotlib.use('TkAgg')
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
    from matplotlib.figure import Figure
    import matplotlib.dates as mdates
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False
    NavigationToolbar2Tk = None


class EmbeddedChart:
    """
    Embedded interactive chart using matplotlib with navigation toolbar.
    """
    
    def __init__(self, parent, width: int = 900, height: int = 600, time_period: str = None):
        self.parent = parent
        self.width = width
        self.height = height
        self.time_period = time_period  # 'hourly', 'daily', 'monthly', or None for dates
        self.fig = None
        self.ax = None
        self.canvas = None
        self.toolbar = None
        self.lines = []  # Store line references for legend toggling
        self.labels = []
        self.original_data = None
        self.current_start_idx = 0
        self.current_end_idx = -1
        
    def create_widget(self) -> tk.Frame:
        """Create and return the chart widget frame."""
        if not MATPLOTLIB_AVAILABLE:
            raise ImportError(
                "matplotlib is required for embedded charts.\n"
                "Install with: pip install matplotlib"
            )
        
        # Create main frame
        self.main_frame = ttk.Frame(self.parent)
        
        # Toolbar frame for date controls
        self.control_frame = ttk.Frame(self.main_frame)
        self.control_frame.pack(fill='x', padx=5, pady=5)
        
        # Date controls
        ttk.Label(self.control_frame, text="From:").pack(side='left', padx=(0, 5))
        self.start_date = ttk.Entry(self.control_frame, width=12)
        self.start_date.pack(side='left', padx=(0, 10))
        
        ttk.Label(self.control_frame, text="To:").pack(side='left', padx=(0, 5))
        self.end_date = ttk.Entry(self.control_frame, width=12)
        self.end_date.pack(side='left', padx=(0, 10))
        
        ttk.Button(self.control_frame, text="Apply", command=self._apply_date_filter).pack(side='left', padx=(0, 5))
        ttk.Button(self.control_frame, text="Reset", command=self._reset_filter).pack(side='left', padx=(0, 15))
        
        # Preset buttons
        presets = [
            ("Today", self._set_today),
            ("7 Days", self._set_7days),
            ("Month", self._set_month),
            ("Year", self._set_year),
            ("All", self._reset_filter)
        ]
        for text, cmd in presets:
            ttk.Button(self.control_frame, text=text, command=cmd).pack(side='left', padx=2)
        
        # Legend toggle frame
        self.legend_frame = ttk.LabelFrame(self.main_frame, text="Legend (Click to Toggle)")
        self.legend_frame.pack(fill='x', padx=5, pady=5)
        self.legend_buttons = []
        
        # Chart frame - use grid for better control
        self.chart_frame = tk.Frame(self.main_frame, bg='#1e293b')
        self.chart_frame.pack(fill='both', expand=True, padx=5, pady=5)
        self.chart_frame.grid_rowconfigure(0, weight=1)
        self.chart_frame.grid_columnconfigure(0, weight=1)
        
        # Create matplotlib figure with dark theme
        self.fig = Figure(figsize=(self.width/100, (self.height-48)/100), dpi=100)
        self.fig.patch.set_facecolor('#1a1a2e')
        
        self.ax = self.fig.add_subplot(111)
        self.ax.set_facecolor('#16213e')
        self.ax.tick_params(colors='white')
        self.ax.xaxis.label.set_color('white')
        self.ax.yaxis.label.set_color('white')
        self.ax.title.set_color('white')
        self.ax.grid(True, alpha=0.3, color='#0f3460')
        
        # Create canvas - use grid instead of pack
        self.canvas = FigureCanvasTkAgg(self.fig, master=self.chart_frame)
        self.canvas.draw()
        self.canvas.get_tk_widget().grid(row=0, column=0, sticky='nsew')
        
        # Create custom navigation toolbar frame below the chart
        self._create_navigation_toolbar()
        
        return self.main_frame
    
    def _create_navigation_toolbar(self):
        """Create a custom-styled navigation toolbar below the graph."""
        # Main toolbar frame with dark background - more prominent
        toolbar_frame = tk.Frame(self.chart_frame, bg='#0f172a', height=48, 
                                  highlightbackground='#334155', highlightthickness=1)
        toolbar_frame.grid(row=1, column=0, sticky='ew')
        toolbar_frame.grid_propagate(False)
        
        # Create the actual matplotlib navigation toolbar (hidden, for functionality)
        self._nav_toolbar = NavigationToolbar2Tk(self.canvas, toolbar_frame, pack_toolbar=False)
        
        # Inner frame for centering content with padding
        inner_frame = tk.Frame(toolbar_frame, bg='#0f172a')
        inner_frame.pack(fill='both', expand=True, padx=8, pady=4)
        
        # Button style configuration - more prominent buttons
        btn_config = {
            'bg': '#3b82f6',  # Brighter blue for visibility
            'fg': 'white',
            'activebackground': '#2563eb',
            'activeforeground': 'white',
            'relief': 'raised',
            'font': ('Segoe UI', 9, 'bold'),
            'padx': 12,
            'pady': 4,
            'cursor': 'hand2',
            'borderwidth': 0,
            'highlightthickness': 0
        }
        
        # Secondary button style (for less prominent actions)
        btn_config_secondary = {
            'bg': '#475569',
            'fg': 'white',
            'activebackground': '#64748b',
            'activeforeground': 'white',
            'relief': 'flat',
            'font': ('Segoe UI', 9),
            'padx': 10,
            'pady': 4,
            'cursor': 'hand2',
            'borderwidth': 0,
            'highlightthickness': 0
        }
        
        # Store references to toolbar buttons
        self.toolbar_buttons = {}
        
        # Helper function to create toolbar buttons
        def create_toolbar_button(parent, text, icon, command, is_primary=True):
            config = btn_config if is_primary else btn_config_secondary
            btn = tk.Button(parent, text=f"{icon}  {text}", command=command, **config)
            btn.pack(side='left', padx=3, pady=2)
            self.toolbar_buttons[text.lower().replace(' ', '_')] = btn
            return btn
        
        # Navigation group label
        nav_label = tk.Label(inner_frame, text="Navigation:", bg='#0f172a', fg='#94a3b8',
                            font=('Segoe UI', 8))
        nav_label.pack(side='left', padx=(0, 5))
        
        # Home button - reset to initial view
        create_toolbar_button(
            inner_frame, "Home", "🏠", 
            lambda: self._nav_home(), is_primary=False
        )
        
        # Back button - previous view
        create_toolbar_button(
            inner_frame, "Back", "◀", 
            lambda: self._nav_back(), is_primary=False
        )
        
        # Forward button - next view
        create_toolbar_button(
            inner_frame, "Forward", "▶", 
            lambda: self._nav_forward(), is_primary=False
        )
        
        # Separator
        tk.Frame(inner_frame, bg='#475569', width=2).pack(side='left', fill='y', padx=10, pady=6)
        
        # Interaction group label
        interact_label = tk.Label(inner_frame, text="Interact:", bg='#0f172a', fg='#94a3b8',
                                 font=('Segoe UI', 8))
        interact_label.pack(side='left', padx=(0, 5))
        
        # Pan button - pan the view (primary)
        self.pan_btn = create_toolbar_button(
            inner_frame, "Pan", "✋", 
            lambda: self._toggle_pan(), is_primary=True
        )
        
        # Zoom button - zoom to rectangle (primary)
        self.zoom_btn = create_toolbar_button(
            inner_frame, "Zoom", "🔍", 
            lambda: self._toggle_zoom(), is_primary=True
        )
        
        # Separator
        tk.Frame(inner_frame, bg='#475569', width=2).pack(side='left', fill='y', padx=10, pady=6)
        
        # Actions group label
        action_label = tk.Label(inner_frame, text="Actions:", bg='#0f172a', fg='#94a3b8',
                               font=('Segoe UI', 8))
        action_label.pack(side='left', padx=(0, 5))
        
        # Save button - save the figure
        create_toolbar_button(
            inner_frame, "Save", "💾", 
            lambda: self._nav_save(), is_primary=False
        )
        
        # Configure/Subplots button
        create_toolbar_button(
            inner_frame, "Configure", "⚙️", 
            lambda: self._nav_configure(), is_primary=False
        )
    
    def _nav_home(self):
        """Reset view to home."""
        self._nav_toolbar.home()
    
    def _nav_back(self):
        """Go back to previous view."""
        self._nav_toolbar.back()
    
    def _nav_forward(self):
        """Go forward to next view."""
        self._nav_toolbar.forward()
    
    def _toggle_pan(self):
        """Toggle pan mode."""
        # Deactivate zoom if active
        if hasattr(self, 'zoom_btn') and self.zoom_btn.cget('relief') == 'sunken':
            self.zoom_btn.config(relief='flat', bg='#334155')
            self._nav_toolbar.zoom()
        
        # Toggle pan
        if self.pan_btn.cget('relief') == 'flat':
            self.pan_btn.config(relief='sunken', bg='#2563eb')
            self._nav_toolbar.pan()
        else:
            self.pan_btn.config(relief='flat', bg='#334155')
            self._nav_toolbar.pan()
    
    def _toggle_zoom(self):
        """Toggle zoom mode."""
        # Deactivate pan if active
        if hasattr(self, 'pan_btn') and self.pan_btn.cget('relief') == 'sunken':
            self.pan_btn.config(relief='flat', bg='#334155')
            self._nav_toolbar.pan()
        
        # Toggle zoom
        if self.zoom_btn.cget('relief') == 'flat':
            self.zoom_btn.config(relief='sunken', bg='#2563eb')
            self._nav_toolbar.zoom()
        else:
            self.zoom_btn.config(relief='flat', bg='#334155')
            self._nav_toolbar.zoom()
    
    def _nav_save(self):
        """Save the figure."""
        self._nav_toolbar.save_figure()
    
    def _nav_configure(self):
        """Open subplots configuration."""
        self._nav_toolbar.configure_subplots()
    
    def load_data(self, 
                  data: List[Dict],
                  title: str = "Interactive Chart",
                  x_title: str = "Time", 
                  y_title: str = "Value",
                  dark_mode: bool = True):
        """Load data and display the chart."""
        self.original_data = data
        self.chart_title = title
        self.x_title = x_title
        self.y_title = y_title
        
        # Clear previous
        self.ax.clear()
        self.lines = []
        self.labels = []
        
        # Apply dark theme
        self.ax.set_facecolor('#16213e')
        self.ax.tick_params(colors='white')
        self.ax.xaxis.label.set_color('white')
        self.ax.yaxis.label.set_color('white')
        self.ax.title.set_color('white')
        self.ax.grid(True, alpha=0.3, color='#0f3460')
        
        # Parse timestamps and values for each series
        # Handle both 'YYYY-MM-DD HH:MM' (hourly) and 'YYYY-MM-DD' (daily/monthly)
        parsed_series = []
        
        for series in data:
            timestamps = []
            values = []
            
            if series.get('timestamps'):
                for ts, val in zip(series['timestamps'], series['values']):
                    try:
                        # Try full datetime format first (hourly data)
                        dt = datetime.strptime(ts, '%Y-%m-%d %H:%M')
                        timestamps.append(dt)
                        values.append(val)
                    except ValueError:
                        try:
                            # Try date format (daily data)
                            dt = datetime.strptime(ts, '%Y-%m-%d')
                            timestamps.append(dt)
                            values.append(val)
                        except ValueError:
                            try:
                                # Try month format (monthly data: "YYYY-MM")
                                dt = datetime.strptime(ts, '%Y-%m')
                                timestamps.append(dt)
                                values.append(val)
                            except ValueError:
                                # Skip invalid timestamps
                                continue
            
            parsed_series.append({
                'name': series.get('name', 'Series'),
                'timestamps': timestamps,
                'values': values,
                'color': series.get('color')
            })
        
        # Check if we should use sequential numbering (1 to N) based on time_period
        use_sequential = self.time_period in ['hourly', 'daily', 'monthly']
        
        # Plot each series
        colors = ['#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', '#ec4899']
        for i, series in enumerate(parsed_series):
            color = series['color'] or colors[i % len(colors)]
            
            if use_sequential and series['timestamps']:
                # Use sequential numbers 1 to N for x-axis
                x_values = list(range(1, len(series['timestamps']) + 1))
                line, = self.ax.plot(x_values, series['values'], 
                                    label=series['name'], color=color, linewidth=1.5)
            else:
                # Use datetime for x-axis
                line, = self.ax.plot(series['timestamps'], series['values'], 
                                    label=series['name'], color=color, linewidth=1.5)
            
            self.lines.append(line)
            self.labels.append(series['name'])
        
        # Set labels and title
        self.ax.set_ylabel(y_title)
        self.ax.set_title(title)
        
        # Format x-axis based on time_period or data density
        if use_sequential:
            # Use sequential numbering: 1 to N
            import matplotlib.ticker as mticker
            if self.time_period == 'hourly':
                self.ax.set_xlabel('Time (hours)')
            elif self.time_period == 'daily':
                self.ax.set_xlabel('Time (days)')
            elif self.time_period == 'monthly':
                self.ax.set_xlabel('Time (months)')
            
            # Use plain integer formatter
            self.ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, p: f'{int(x)}'))
            # Set reasonable number of ticks
            self.ax.xaxis.set_major_locator(mticker.MaxNLocator(10))
        else:
            # Use date formatting
            self.ax.set_xlabel(x_title)
            import matplotlib.dates as mdates
            
            # Determine appropriate date format based on data range
            if timestamps:
                time_range = timestamps[-1] - timestamps[0]
                days_range = time_range.days
                
                if days_range <= 1:
                    # Less than 1 day - show hours
                    self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%m/%d %H:%M'))
                    self.ax.xaxis.set_major_locator(mdates.HourLocator(interval=3))
                elif days_range <= 31:
                    # Up to 1 month - show days
                    self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%m/%d'))
                    self.ax.xaxis.set_major_locator(mdates.DayLocator(interval=max(1, days_range // 10)))
                elif days_range <= 365:
                    # Up to 1 year - show months
                    self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %Y'))
                    self.ax.xaxis.set_major_locator(mdates.MonthLocator())
                else:
                    # More than 1 year - show years
                    self.ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y'))
                    self.ax.xaxis.set_major_locator(mdates.YearLocator())
        
        # Keep labels horizontal (no rotation)
        for label in self.ax.get_xticklabels():
            label.set_rotation(0)
            label.set_horizontalalignment('center')
        
        # Add legend
        self.ax.legend(loc='upper left', facecolor='#1e293b', edgecolor='#334155', 
                      labelcolor='white')
        
        # Create custom legend toggle buttons
        self._create_legend_buttons()
        
        # Set date inputs
        if timestamps:
            self.start_date.delete(0, 'end')
            self.start_date.insert(0, timestamps[0].strftime('%Y-%m-%d'))
            self.end_date.delete(0, 'end')
            self.end_date.insert(0, timestamps[-1].strftime('%Y-%m-%d'))
        
        # Add tooltip annotation
        self.tooltip = self.ax.annotate('', xy=(0, 0), xytext=(10, 10),
                                       textcoords='offset points',
                                       bbox=dict(boxstyle='round', facecolor='#1e293b', 
                                                edgecolor='#3b82f6', alpha=0.9),
                                       color='white', fontsize=9,
                                       visible=False)
        
        # Connect hover event for tooltips
        self.canvas.mpl_connect('motion_notify_event', self._on_hover)
        
        self.fig.tight_layout()
        self.canvas.draw()
    
    def _create_legend_buttons(self):
        """Create clickable legend toggle buttons."""
        # Clear old buttons
        for btn in self.legend_buttons:
            btn.destroy()
        self.legend_buttons = []
        
        # Create new toggle buttons
        for i, (line, label) in enumerate(zip(self.lines, self.labels)):
            btn = tk.Button(self.legend_frame, text=f"👁 {label}", 
                          command=lambda idx=i: self._toggle_series(idx),
                          bg='#3b82f6', fg='white', relief='raised',
                          font=('Segoe UI', 9), padx=10, pady=2)
            btn.pack(side='left', padx=5, pady=5)
            self.legend_buttons.append(btn)
    
    def _toggle_series(self, index):
        """Toggle visibility of a series."""
        line = self.lines[index]
        btn = self.legend_buttons[index]
        
        if line.get_visible():
            line.set_visible(False)
            btn.config(bg='#6b7280', relief='sunken')
        else:
            line.set_visible(True)
            btn.config(bg='#3b82f6', relief='raised')
        
        self.canvas.draw()
    
    def _on_hover(self, event):
        """Show tooltip on hover."""
        if event.inaxes != self.ax:
            self.tooltip.set_visible(False)
            self.canvas.draw_idle()
            return
        
        # Find nearest point
        min_dist = float('inf')
        nearest_info = None
        
        # Convert matplotlib date (float) to datetime for comparison
        import matplotlib.dates as mdates
        
        for line in self.lines:
            if not line.get_visible():
                continue
            
            xdata = line.get_xdata()
            ydata = line.get_ydata()
            
            for i, (x, y) in enumerate(zip(xdata, ydata)):
                # x is already a datetime object (we stored them)
                if isinstance(x, datetime):
                    # Convert to matplotlib date for comparison
                    x_num = mdates.date2num(x)
                else:
                    x_num = x
                
                dist = abs(event.xdata - x_num) if event.xdata else float('inf')
                if dist < min_dist and dist < 10:  # Within 10 days
                    min_dist = dist
                    nearest_info = (x, y, line.get_label())
        
        if nearest_info:
            x, y, label = nearest_info
            if isinstance(x, datetime):
                # Format based on data granularity
                if x.day == 1 and x.hour == 0 and x.minute == 0:
                    # Monthly data (first day of month, midnight)
                    x_str = x.strftime('%Y-%m')
                elif x.hour == 0 and x.minute == 0:
                    # Daily data (midnight)
                    x_str = x.strftime('%Y-%m-%d')
                else:
                    # Hourly data
                    x_str = x.strftime('%Y-%m-%d %H:%M')
            else:
                x_str = str(x)
            
            self.tooltip.xy = (event.xdata, event.ydata)
            self.tooltip.set_text(f"{label}\nTime: {x_str}\nValue: {y:.3f}")
            self.tooltip.set_visible(True)
            self.canvas.draw_idle()
        else:
            self.tooltip.set_visible(False)
            self.canvas.draw_idle()
    
    def _apply_date_filter(self):
        """Apply date range filter."""
        try:
            start_str = self.start_date.get()
            end_str = self.end_date.get()
            
            start = datetime.strptime(start_str, '%Y-%m-%d')
            end = datetime.strptime(end_str, '%Y-%m-%d')
            end = end.replace(hour=23, minute=59, second=59)
            
            # Filter data
            filtered_data = []
            for series in self.original_data:
                filtered_timestamps = []
                filtered_values = []
                
                for ts_str, val in zip(series['timestamps'], series['values']):
                    # Parse timestamp - handle hourly, daily, and monthly formats
                    try:
                        ts = datetime.strptime(ts_str, '%Y-%m-%d %H:%M')
                    except ValueError:
                        try:
                            ts = datetime.strptime(ts_str, '%Y-%m-%d')
                        except ValueError:
                            try:
                                ts = datetime.strptime(ts_str, '%Y-%m')
                            except ValueError:
                                continue
                    if start <= ts <= end:
                        filtered_timestamps.append(ts_str)
                        filtered_values.append(val)
                
                if filtered_timestamps:
                    filtered_series = series.copy()
                    filtered_series['timestamps'] = filtered_timestamps
                    filtered_series['values'] = filtered_values
                    filtered_data.append(filtered_series)
            
            if filtered_data:
                self.load_data(filtered_data, self.chart_title, self.x_title, self.y_title)
        except ValueError as e:
            print(f"Date filter error: {e}")
    
    def _reset_filter(self):
        """Reset to show all data."""
        if self.original_data:
            self.load_data(self.original_data, self.chart_title, self.x_title, self.y_title)
    
    def _set_today(self):
        """Set filter to today."""
        today = datetime.now()
        self.start_date.delete(0, 'end')
        self.start_date.insert(0, today.strftime('%Y-%m-%d'))
        self.end_date.delete(0, 'end')
        self.end_date.insert(0, today.strftime('%Y-%m-%d'))
        self._apply_date_filter()
    
    def _set_7days(self):
        """Set filter to last 7 days."""
        end = datetime.now()
        start = end - timedelta(days=7)
        self.start_date.delete(0, 'end')
        self.start_date.insert(0, start.strftime('%Y-%m-%d'))
        self.end_date.delete(0, 'end')
        self.end_date.insert(0, end.strftime('%Y-%m-%d'))
        self._apply_date_filter()
    
    def _set_month(self):
        """Set filter to last month."""
        end = datetime.now()
        start = end - timedelta(days=30)
        self.start_date.delete(0, 'end')
        self.start_date.insert(0, start.strftime('%Y-%m-%d'))
        self.end_date.delete(0, 'end')
        self.end_date.insert(0, end.strftime('%Y-%m-%d'))
        self._apply_date_filter()
    
    def _set_year(self):
        """Set filter to last year."""
        end = datetime.now()
        start = end - timedelta(days=365)
        self.start_date.delete(0, 'end')
        self.start_date.insert(0, start.strftime('%Y-%m-%d'))
        self.end_date.delete(0, 'end')
        self.end_date.insert(0, end.strftime('%Y-%m-%d'))
        self._apply_date_filter()
    
    def destroy(self):
        """Cleanup resources."""
        if self.fig:
            import matplotlib.pyplot as plt
            plt.close(self.fig)


def create_embedded_chart(parent, 
                         hourly_results: List[Dict],
                         selected_model: str,
                         graph_type: str,
                         width: int = 900,
                         height: int = 600,
                         dark_mode: bool = True,
                         time_period: str = None) -> Tuple[EmbeddedChart, tk.Frame]:
    """
    Create and return an embedded chart widget.
    
    Args:
        time_period: 'hourly', 'daily', 'monthly', or None for date-based x-axis
    
    Example:
        chart, widget = create_embedded_chart(
            parent_frame,
            hourly_results,
            "Model 1",
            "Wind Speed",
            time_period="hourly"
        )
        widget.pack(fill="both", expand=True)
    """
    from interactive_chart import ChartDataBuilder
    
    # Extract data
    builder = ChartDataBuilder()
    timestamps, values, title, y_label = builder.from_hourly_results(
        hourly_results, selected_model, graph_type
    )
    
    # Create embedded chart with time_period for sequential x-axis numbering
    chart = EmbeddedChart(parent, width=width, height=height, time_period=time_period)
    widget = chart.create_widget()
    
    # Define colors for each graph type
    color_map = {
        "Temperature": "#ef4444",      # 🔴 Red
        "Wind": "#3b82f6",             # 🔵 Blue
        "V (Wind Speed)": "#3b82f6",   # 🔵 Blue
        "Air Density": "#10b981",      # 🟢 Green
        "ρ/ρ₀": "#8b5cf6",             # 🟣 Purple (Air Density Ratio)
        "Power (kW)": "#f59e0b",       # 🟡 Yellow
        "Energy (kWh)": "#f59e0b",     # 🟡 Yellow
        "Power P (kW)": "#f59e0b",     # 🟡 Yellow
        "Energy E (kWh)": "#f59e0b",   # 🟡 Yellow
    }
    
    # Get color for this graph type, default to blue
    line_color = color_map.get(graph_type, "#3b82f6")
    
    # Prepare data series
    data = [{
        "name": graph_type,
        "timestamps": timestamps,
        "values": values,
        "color": line_color
    }]
    
    # Load data
    chart.load_data(data, title, "Time", y_label, dark_mode)
    
    return chart, widget
