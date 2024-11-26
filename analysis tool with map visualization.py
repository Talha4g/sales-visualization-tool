import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import numpy as np
import pandas as pd
import io
import folium
from folium import plugins
import webbrowser
import os

class SalesVisualizationTool:
    def __init__(self, root):
        self.root = root
        self.root.title("Sales Analysis Tool")
        self.root.geometry("1200x900")
        
        # Store the full dataset
        self.df = pd.DataFrame()
        
        # Create main container
        self.main_container = ttk.Notebook(root)
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Create tabs
        self.data_tab = ttk.Frame(self.main_container)
        self.analysis_tab = ttk.Frame(self.main_container)
        self.visualization_tab = ttk.Frame(self.main_container)
        self.map_tab = ttk.Frame(self.main_container)
        
        self.main_container.add(self.data_tab, text='Data Import')
        self.main_container.add(self.analysis_tab, text='Analysis')
        self.main_container.add(self.visualization_tab, text='Visualizations')
        self.main_container.add(self.map_tab, text='Map View')
        
        self._create_data_tab()
        self._create_analysis_tab()
        self._create_visualization_tab()
        self._create_map_tab()

    def _create_data_tab(self):
        # Analysis Type Selection Frame
        analysis_frame = ttk.LabelFrame(self.data_tab, text="Select Analysis Type", padding="10")
        analysis_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Analysis Type Radio Buttons
        self.analysis_type = tk.StringVar(value="yoy")
        ttk.Radiobutton(analysis_frame, text="Year over Year Analysis", 
                       variable=self.analysis_type, value="yoy",
                       command=self._update_help_text).pack(side=tk.LEFT, padx=20)
        ttk.Radiobutton(analysis_frame, text="Sales vs Target Analysis", 
                       variable=self.analysis_type, value="target",
                       command=self._update_help_text).pack(side=tk.LEFT, padx=20)
        ttk.Radiobutton(analysis_frame, text="Map Analysis", 
                       variable=self.analysis_type, value="map",
                       command=self._update_help_text).pack(side=tk.LEFT, padx=20)
        
        # Import frame
        import_frame = ttk.LabelFrame(self.data_tab, text="Import Data", padding="10")
        import_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(import_frame, text="Import Excel File", 
                  command=self._import_excel_file).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(import_frame, text="Or paste Excel data:").pack(side=tk.LEFT, padx=5)
        self.paste_area = tk.Text(import_frame, height=4, width=50)
        self.paste_area.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(import_frame, text="Import Pasted Data", 
                  command=self._import_pasted_data).pack(side=tk.LEFT, padx=5)
        
        # Dynamic help text based on analysis type
        self.help_label = ttk.Label(import_frame, foreground="gray")
        self.help_label.pack(side=tk.LEFT, padx=5)
        
        # Table frame
        table_frame = ttk.LabelFrame(self.data_tab, text="Data Preview", padding="10")
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create Treeview with scrollbar
        self.tree_scroll = ttk.Scrollbar(table_frame)
        self.tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Create Treeview with initial columns
        if self.analysis_type.get() == "yoy":
            columns = ('customer', 'current_cases', 'previous_cases', 'growth')
        elif self.analysis_type.get() == "target":
            columns = ('customer', 'current_cases', 'target', 'achievement')
        else:  # map analysis
            columns = ('customer', 'sales', 'latitude', 'longitude')
            
        self.tree = ttk.Treeview(table_frame, columns=columns, show='headings',
                                yscrollcommand=self.tree_scroll.set)
        
        # Set initial column headings
        self._update_tree_columns()
        
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.tree_scroll.config(command=self.tree.yview)
        
        # Update help text
        self._update_help_text()

    def _create_map_tab(self):
        # Control Frame
        control_frame = ttk.LabelFrame(self.map_tab, text="Map Controls", padding="10")
        control_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(control_frame, text="Generate Map", 
                  command=self._generate_map).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(control_frame, text="Open in Browser", 
                  command=self._open_map_in_browser).pack(side=tk.LEFT, padx=5)
        
        # Map info frame
        self.map_info_frame = ttk.LabelFrame(self.map_tab, text="Map Information", padding="10")
        self.map_info_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(self.map_info_frame, 
                 text="Hover over points to see customer and sales information").pack()

    def _generate_map(self):
        try:
            if len(self.df) == 0:
                messagebox.showwarning("Warning", "Please import data first!")
                return
                
            # Create map centered on the mean of coordinates
            center_lat = self.df['Latitude'].mean()
            center_lon = self.df['Longitude'].mean()
            m = folium.Map(location=[center_lat, center_lon], zoom_start=4)
            
            # Add markers for each point
            for idx, row in self.df.iterrows():
                popup_content = f"""
                    <div style='width:200px'>
                        <b>Customer:</b> {row['Customer']}<br>
                        <b>Sales:</b> {row['Sales']:,.0f}<br>
                        <b>Location:</b> {row['Latitude']:.4f}, {row['Longitude']:.4f}
                    </div>
                """
                
                folium.CircleMarker(
                    location=[row['Latitude'], row['Longitude']],
                    radius=np.sqrt(row['Sales'])/100,  # Size based on sales
                    popup=folium.Popup(popup_content, max_width=300),
                    tooltip=f"{row['Customer']}: {row['Sales']:,.0f}",
                    color='blue',
                    fill=True,
                    fill_opacity=0.6
                ).add_to(m)
            
            # Add heatmap layer
            heat_data = [[row['Latitude'], row['Longitude'], row['Sales']] 
                        for idx, row in self.df.iterrows()]
            plugins.HeatMap(heat_data).add_to(m)
            
            # Save map
            self.map_path = "sales_map.html"
            m.save(self.map_path)
            
            # Open in default browser
            webbrowser.open(self.map_path)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error generating map: {str(e)}")

    def _open_map_in_browser(self):
        if hasattr(self, 'map_path') and os.path.exists(self.map_path):
            webbrowser.open(self.map_path)
        else:
            messagebox.showwarning("Warning", "Please generate the map first!")

    def _update_help_text(self):
        if self.analysis_type.get() == "yoy":
            help_text = "Column order: 1st=Customer, 2nd=Current Sales, 3rd=Previous Sales"
        elif self.analysis_type.get() == "target":
            help_text = "Column order: 1st=Customer, 2nd=Current Sales, 3rd=Target"
        else:  # map analysis
            help_text = "Column order: 1st=Customer, 2nd=Sales, 3rd=Latitude, 4th=Longitude"
        self.help_label.config(text=help_text)
        self._update_tree_columns()

    def _update_tree_columns(self):
        if self.analysis_type.get() == "yoy":
            columns = ('customer', 'current_cases', 'previous_cases', 'growth')
            self.tree['columns'] = columns
            
            self.tree.heading('customer', text='Customer')
            self.tree.heading('current_cases', text='Current Cases')
            self.tree.heading('previous_cases', text='Previous Cases')
            self.tree.heading('growth', text='Growth %')
        elif self.analysis_type.get() == "target":
            columns = ('customer', 'current_cases', 'target', 'achievement')
            self.tree['columns'] = columns
            
            self.tree.heading('customer', text='Customer')
            self.tree.heading('current_cases', text='Current Cases')
            self.tree.heading('target', text='Target Cases')
            self.tree.heading('achievement', text='Achievement %')
        else:  # map analysis
            columns = ('customer', 'sales', 'latitude', 'longitude')
            self.tree['columns'] = columns
            
            self.tree.heading('customer', text='Customer')
            self.tree.heading('sales', text='Sales')
            self.tree.heading('latitude', text='Latitude')
            self.tree.heading('longitude', text='Longitude')
        
        # Set column widths
        for col in columns:
            self.tree.column(col, width=150)

    def _create_analysis_tab(self):
        # Control Frame
        control_frame = ttk.LabelFrame(self.analysis_tab, text="Analysis Controls", padding="10")
        control_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Top N control
        ttk.Label(control_frame, text="Show Top N:").pack(side=tk.LEFT, padx=5)
        self.top_n_var = tk.StringVar(value="10")
        top_n_entry = ttk.Entry(control_frame, textvariable=self.top_n_var, width=5)
        top_n_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(control_frame, text="Update Analysis", 
                  command=self._update_analysis).pack(side=tk.LEFT, padx=5)
        
        # Summary Frame
        self.summary_frame = ttk.LabelFrame(self.analysis_tab, text="Summary Statistics", padding="10")
        self.summary_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Analysis Plot Frame
        self.analysis_plot_frame = ttk.Frame(self.analysis_tab)
        self.analysis_plot_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create figure and canvas for analysis
        self.analysis_fig = Figure(figsize=(10, 6))
        self.analysis_canvas = FigureCanvasTkAgg(self.analysis_fig, master=self.analysis_plot_frame)
        self.analysis_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # Add toolbar
        self.analysis_toolbar = NavigationToolbar2Tk(self.analysis_canvas, self.analysis_plot_frame)
        self.analysis_toolbar.update()

    def _create_visualization_tab(self):
        # Control Frame
        control_frame = ttk.Frame(self.visualization_tab)
        control_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Chart type selector
        ttk.Label(control_frame, text="Chart Type:").pack(side=tk.LEFT, padx=5)
        self.chart_type = tk.StringVar(value="bar")
        chart_combo = ttk.Combobox(control_frame, textvariable=self.chart_type,
                                 values=["bar", "line", "scatter", "pie"])
        chart_combo.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(control_frame, text="Update Visualization", 
                  command=self._update_visualization).pack(side=tk.LEFT, padx=5)
        
        # Visualization Plot Frame
        self.viz_plot_frame = ttk.Frame(self.visualization_tab)
        self.viz_plot_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create figure and canvas for visualization
        self.viz_fig = Figure(figsize=(10, 6))
        self.viz_canvas = FigureCanvasTkAgg(self.viz_fig, master=self.viz_plot_frame)
        self.viz_canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # Add toolbar
        self.viz_toolbar = NavigationToolbar2Tk(self.viz_canvas, self.viz_plot_frame)
        self.viz_toolbar.update()

    def _import_excel_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if file_path:
            try:
                self.df = pd.read_excel(file_path)
                
                # Rename columns based on analysis type and position
                if self.analysis_type.get() == "map":
                    if len(self.df.columns) >= 4:
                        new_columns = {
                            self.df.columns[0]: 'Customer',
                            self.df.columns[1]: 'Sales',
                            self.df.columns[2]: 'Latitude',
                            self.df.columns[3]: 'Longitude'
                        }
                        self.df = self.df.rename(columns=new_columns)
                        self.df = self.df[['Customer', 'Sales', 'Latitude', 'Longitude']]
                    else:
                        messagebox.showerror("Error", "Excel file must have at least 4 columns for map analysis")
                        return
                elif self.analysis_type.get() == "yoy":
                    if len(self.df.columns) >= 3:
                        new_columns = {
                            self.df.columns[0]: 'Customer',
                            self.df.columns[1]: 'Current Sales',
                            self.df.columns[2]: 'Previous Sales'
                        }
                        self.df = self.df.rename(columns=new_columns)
                        self.df = self.df[['Customer', 'Current Sales', 'Previous Sales']]
                    else:
                        messagebox.showerror("Error", "Excel file must have at least 3 columns for YoY analysis")
                        return
                else:  # target analysis
                    if len(self.df.columns) >= 3:
                        new_columns = {
                            self.df.columns[0]: 'Customer',
                            self.df.columns[1]: 'Current Sales',
                            self.df.columns[2]: 'Target'
                        }
                        self.df = self.df.rename(columns=new_columns)
                        self.df = self.df[['Customer', 'Current Sales', 'Target']]
                    else:
                        messagebox.showerror("Error", "Excel file must have at least 3 columns for target analysis")
                        return
                
                self._process_data()
            except Exception as e:
                messagebox.showerror("Error", f"Error importing Excel file: {str(e)}")

    def _import_pasted_data(self):
        try:
            data = self.paste_area.get("1.0", tk.END).strip()
            if not data:
                messagebox.showwarning("Warning", "Please paste some data first!")
                return
            
            # Try different separators
            try:
                # First try tab-separated
                self.df = pd.read_csv(io.StringIO(data), sep='\t')
            except:
                try:
                    # Then try comma-separated
                    self.df = pd.read_csv(io.StringIO(data), sep=',')
                except:
                    messagebox.showerror("Error", "Could not parse the pasted data. Please check the format.")
                    return
            
            # Rename columns based on analysis type and position
            if self.analysis_type.get() == "map":
                if len(self.df.columns) >= 4:
                    new_columns = {
                        self.df.columns[0]: 'Customer',
                        self.df.columns[1]: 'Sales',
                        self.df.columns[2]: 'Latitude',
                        self.df.columns[3]: 'Longitude'
                    }
                    self.df = self.df.rename(columns=new_columns)
                    self.df = self.df[['Customer', 'Sales', 'Latitude', 'Longitude']]
                else:
                    messagebox.showerror("Error", "Please paste data with at least 4 columns for map analysis")
                    return
            elif self.analysis_type.get() == "yoy":
                if len(self.df.columns) >= 3:
                    new_columns = {
                        self.df.columns[0]: 'Customer',
                        self.df.columns[1]: 'Current Sales',
                        self.df.columns[2]: 'Previous Sales'
                    }
                    self.df = self.df.rename(columns=new_columns)
                    self.df = self.df[['Customer', 'Current Sales', 'Previous Sales']]
                else:
                    messagebox.showerror("Error", "Please paste data with at least 3 columns for YoY analysis")
                    return
            else:  # target analysis
                if len(self.df.columns) >= 3:
                    new_columns = {
                        self.df.columns[0]: 'Customer',
                        self.df.columns[1]: 'Current Sales',
                        self.df.columns[2]: 'Target'
                    }
                    self.df = self.df.rename(columns=new_columns)
                    self.df = self.df[['Customer', 'Current Sales', 'Target']]
                else:
                    messagebox.showerror("Error", "Please paste data with at least 3 columns for target analysis")
                    return
            
            self._process_data()
            self.paste_area.delete("1.0", tk.END)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error processing pasted data: {str(e)}")

    def _process_data(self):
        try:
            analysis_type = self.analysis_type.get()
            
            if analysis_type == "map":
                # Convert numeric columns for map analysis
                self.df['Sales'] = pd.to_numeric(self.df['Sales'], errors='coerce')
                self.df['Latitude'] = pd.to_numeric(self.df['Latitude'], errors='coerce')
                self.df['Longitude'] = pd.to_numeric(self.df['Longitude'], errors='coerce')
            elif analysis_type == "yoy":
                # Convert and calculate for YoY analysis
                self.df['Current Sales'] = pd.to_numeric(self.df['Current Sales'], errors='coerce')
                self.df['Previous Sales'] = pd.to_numeric(self.df['Previous Sales'], errors='coerce')
                self.df['Growth'] = ((self.df['Current Sales'] - self.df['Previous Sales']) / 
                                   self.df['Previous Sales'] * 100)
            else:  # target analysis
                # Convert and calculate for target analysis
                self.df['Current Sales'] = pd.to_numeric(self.df['Current Sales'], errors='coerce')
                self.df['Target'] = pd.to_numeric(self.df['Target'], errors='coerce')
                self.df['Achievement'] = (self.df['Current Sales'] / self.df['Target'] * 100)
            
            # Drop any rows with NaN values after conversion
            self.df = self.df.dropna()
            
            # Update displays
            self._update_tree()
            if analysis_type != "map":
                self._update_analysis()
                self._update_visualization()
            
            messagebox.showinfo("Success", f"Imported {len(self.df)} records successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error processing data: {str(e)}")

    def _update_tree(self):
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Add new items based on analysis type
        if self.analysis_type.get() == "map":
            for _, row in self.df.iterrows():
                self.tree.insert('', tk.END, values=(
                    row['Customer'],
                    f"{row['Sales']:,.0f}",
                    f"{row['Latitude']:.4f}",
                    f"{row['Longitude']:.4f}"
                ))
        elif self.analysis_type.get() == "yoy":
            for _, row in self.df.iterrows():
                self.tree.insert('', tk.END, values=(
                    row['Customer'],
                    f"{row['Current Sales']:,.0f}",
                    f"{row['Previous Sales']:,.0f}",
                    f"{row['Growth']:,.1f}%"
                ))
        else:  # target analysis
            for _, row in self.df.iterrows():
                self.tree.insert('', tk.END, values=(
                    row['Customer'],
                    f"{row['Current Sales']:,.0f}",
                    f"{row['Target']:,.0f}",
                    f"{row['Achievement']:,.1f}%"
                ))

    def _update_analysis(self):
        try:
            # Clear previous summary
            for widget in self.summary_frame.winfo_children():
                widget.destroy()
            
            analysis_type = self.analysis_type.get()
            
            if analysis_type == "yoy":
                # Year over Year Analysis
                summary_stats = {
                    'Total Customers': len(self.df),
                    'Total Current Cases': self.df['Current Sales'].sum(),
                    'Total Previous Cases': self.df['Previous Sales'].sum(),
                    'Average Growth': self.df['Growth'].mean(),
                    'Median Growth': self.df['Growth'].median(),
                    'Customers with Positive Growth': (self.df['Growth'] > 0).sum(),
                    'Customers with Negative Growth': (self.df['Growth'] < 0).sum()
                }
            elif analysis_type == "target":
                # Target Analysis
                summary_stats = {
                    'Total Customers': len(self.df),
                    'Total Current Cases': self.df['Current Sales'].sum(),
                    'Total Target': self.df['Target'].sum(),
                    'Overall Achievement': (self.df['Current Sales'].sum() / self.df['Target'].sum() * 100),
                    'Average Achievement': self.df['Achievement'].mean(),
                    'Customers Above Target': (self.df['Achievement'] >= 100).sum(),
                    'Customers Below Target': (self.df['Achievement'] < 100).sum()
                }
            else:  # map analysis
                return  # No summary stats for map view
            
            # Display summary statistics
            for i, (key, value) in enumerate(summary_stats.items()):
                if 'Cases' in key or 'Target' in key:
                    value_str = f"{value:,.0f}"
                elif 'Growth' in key or 'Achievement' in key:
                    value_str = f"{value:.1f}%"
                else:
                    value_str = str(value)
                    
                ttk.Label(self.summary_frame, text=f"{key}:").grid(row=i//3, column=(i%3)*2, padx=5, pady=2)
                ttk.Label(self.summary_frame, text=value_str).grid(row=i//3, column=(i%3)*2+1, padx=5, pady=2)
            
            # Update plot
            self.analysis_fig.clear()
            ax = self.analysis_fig.add_subplot(111)
            
            try:
                top_n = int(self.top_n_var.get())
            except:
                top_n = 10
            
            if analysis_type == "yoy":
                # Year over Year comparison
                top_customers = self.df.nlargest(top_n, 'Current Sales')
                
                x = np.arange(len(top_customers))
                width = 0.35
                
                ax.bar(x - width/2, top_customers['Current Sales'], width, label='Current Year')
                ax.bar(x + width/2, top_customers['Previous Sales'], width, label='Previous Year')
                
                ax.set_ylabel('Number of Cases')
                ax.set_title(f'Top {top_n} Customers - Year over Year Comparison')
                
            elif analysis_type == "target":
                # Target comparison - showing Actual vs Target
                top_customers = self.df.nlargest(top_n, 'Current Sales')
                
                x = np.arange(len(top_customers))
                width = 0.35
                
                ax.bar(x - width/2, top_customers['Current Sales'], width, label='Actual')
                ax.bar(x + width/2, top_customers['Target'], width, label='Target')
                
                ax.set_ylabel('Number of Cases')
                ax.set_title(f'Top {top_n} Customers - Actual vs Target')
            
            if analysis_type in ["yoy", "target"]:
                ax.set_xticks(x)
                ax.set_xticklabels(top_customers['Customer'], rotation=45, ha='right')
                ax.legend()
                
                # Format y-axis to show whole numbers with commas
                ax.yaxis.set_major_formatter(lambda x, p: f'{int(x):,}')
                
                self.analysis_fig.tight_layout()
                self.analysis_canvas.draw()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error updating analysis: {str(e)}")

    def _update_visualization(self):
        try:
            self.viz_fig.clear()
            ax = self.viz_fig.add_subplot(111)
            
            chart_type = self.chart_type.get()
            analysis_type = self.analysis_type.get()
            
            if analysis_type == "map":
                return  # No additional visualizations for map view
            
            if chart_type == "bar":
                if analysis_type == "yoy":
                    # Sort by growth
                    sorted_df = self.df.sort_values('Growth', ascending=False)
                    ax.bar(range(len(sorted_df)), sorted_df['Growth'])
                    ax.set_title('Customer Growth Distribution')
                    ax.set_xlabel('Customers (sorted by growth)')
                    ax.set_ylabel('Growth (%)')
                else:
                    # Sort by achievement
                    sorted_df = self.df.sort_values('Achievement', ascending=False)
                    ax.bar(range(len(sorted_df)), sorted_df['Achievement'])
                    ax.set_title('Customer Achievement Distribution')
                    ax.set_xlabel('Customers (sorted by achievement)')
                    ax.set_ylabel('Achievement (%)')
                
            elif chart_type == "line":
                if analysis_type == "yoy":
                    # Growth trend
                    sorted_df = self.df.sort_values('Growth')
                    ax.plot(range(len(sorted_df)), sorted_df['Growth'])
                    ax.set_title('Growth Trend')
                    ax.set_xlabel('Customers (sorted by growth)')
                    ax.set_ylabel('Growth (%)')
                else:
                    # Achievement trend
                    sorted_df = self.df.sort_values('Achievement')
                    ax.plot(range(len(sorted_df)), sorted_df['Achievement'])
                    ax.set_title('Achievement Trend')
                    ax.set_xlabel('Customers (sorted by achievement)')
                    ax.set_ylabel('Achievement (%)')
                
            elif chart_type == "scatter":
                if analysis_type == "yoy":
                    ax.scatter(self.df['Previous Sales'], self.df['Current Sales'])
                    ax.set_title('Current vs Previous Cases')
                    ax.set_xlabel('Previous Cases')
                    ax.set_ylabel('Current Cases')
                else:
                    ax.scatter(self.df['Target'], self.df['Current Sales'])
                    ax.set_title('Actual vs Target Cases')
                    ax.set_xlabel('Target Cases')
                    ax.set_ylabel('Actual Cases')
                
                # Add diagonal line for reference
                max_val = max(
                    self.df['Current Sales'].max(),
                    self.df['Previous Sales' if analysis_type == "yoy" else 'Target'].max()
                )
                ax.plot([0, max_val], [0, max_val], 'r--', alpha=0.5)
                
            elif chart_type == "pie":
                if analysis_type == "yoy":
                    # Group by growth ranges
                    bins = [-float('inf'), -10, 0, 10, float('inf')]
                    labels = ['High Decline (<-10%)', 'Slight Decline (-10-0%)', 
                             'Slight Growth (0-10%)', 'High Growth (>10%)']
                    growth_cats = pd.cut(self.df['Growth'], bins=bins, labels=labels)
                    growth_dist = growth_cats.value_counts()
                    
                    ax.pie(growth_dist, labels=growth_dist.index, autopct='%1.1f%%')
                    ax.set_title('Distribution of Customer Growth')
                else:
                    # Group by achievement ranges
                    bins = [-float('inf'), 80, 90, 100, 110, float('inf')]
                    labels = ['Below 80%', '80-90%', '90-100%', '100-110%', 'Above 110%']
                    achievement_cats = pd.cut(self.df['Achievement'], bins=bins, labels=labels)
                    achievement_dist = achievement_cats.value_counts()
                    
                    ax.pie(achievement_dist, labels=achievement_dist.index, autopct='%1.1f%%')
                    ax.set_title('Distribution of Target Achievement')
            
            self.viz_fig.tight_layout()
            self.viz_canvas.draw()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error updating visualization: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = SalesVisualizationTool(root)
    root.mainloop()