"""
Interactive Chart Data Builder
Helper module for building chart data from energy calculation results.
"""

from typing import List, Dict, Tuple


class ChartDataBuilder:
    """Helper class to build chart data from energy calculation results."""
    
    @staticmethod
    def from_hourly_results(hourly_results: List[Dict], 
                           selected_model: str,
                           graph_type: str) -> Tuple[List[str], List[float], str, str]:
        """
        Extract timestamps and values from results based on graph type.
        Handles both hourly data (with 'timestamp') and aggregated data (with 'period').
        
        Returns: (timestamps, values, title, y_label)
        """
        timestamps = []
        values = []
        
        for row in hourly_results:
            # Handle both 'timestamp' (hourly) and 'period' (daily/monthly) keys
            ts = row.get("timestamp") or row.get("period", "")
            if not ts:
                continue
            timestamps.append(ts)
            
            if "Temperature" in graph_type:
                # For aggregated data, use avg_temp_c instead of temp_c
                val = row.get("temp_c") or row.get("avg_temp_c", 0)
                values.append(val)
                y_label = "Temperature (°C)"
                title = f"Temperature - {selected_model}"
            elif "Wind" in graph_type or graph_type == "V (Wind Speed)":
                val = row.get("wind_speed") or row.get("avg_wind_speed", 0)
                values.append(val)
                y_label = "Wind Speed (m/s)"
                title = f"Wind Speed - {selected_model}"
            elif "Air Density" in graph_type:
                val = row.get("air_density") or row.get("avg_air_density", 0)
                values.append(val)
                y_label = "Air Density (kg/m³)"
                title = f"Air Density - {selected_model}"
            elif "ρ/ρ₀" in graph_type or "rho_ratio" in graph_type:
                val = row.get("rho_ratio") or row.get("avg_rho_ratio", 1.0)
                values.append(val)
                y_label = "ρ/ρ₀ (Air Density Ratio)"
                title = f"Air Density Ratio - {selected_model}"
            elif "Power" in graph_type or "Energy" in graph_type or "P (kW)" in graph_type or "E (kWh)" in graph_type:
                val = row.get("power_kw") or row.get("total_energy_kwh", 0)
                values.append(val)
                y_label = "Power (kW)"
                title = f"Power - {selected_model}"
            else:
                values.append(0)
                y_label = "Value"
                title = f"Chart - {selected_model}"
                
        return timestamps, values, title, y_label
