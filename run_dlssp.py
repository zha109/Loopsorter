import pandas as pd
from datetime import datetime, timedelta
import math

Umax = 0.85
beta = 0.5
gamma = 1.0
delta = 1.0
theta = 0.9

def compute_travel_time(quantity, lane_speed, distance_factor=1.0):
    return timedelta(minutes=quantity * distance_factor / lane_speed * 10)

def compute_completion_time(start, travel, processing, packing, induction):
    return start + travel + timedelta(minutes=processing) + timedelta(minutes=packing) + induction

def compute_induction_time(sku):
    return timedelta(minutes=1 + (hash(sku) % 3))

def decode(orders_df):
    results = []
    lane_last_end = {}
    lane_total_time = {}
    current_time = datetime.now().replace(hour=8, minute=0, second=0, microsecond=0)
    orders_sorted = orders_df.sort_values(by='ReleaseTime')
    active_trays = 0
    for _, row in orders_sorted.iterrows():
        lane = row['Lane']
        release = current_time + timedelta(minutes=row['ReleaseTime'])
        lane_speed = row['LaneSpeed']
        processing_time = row.get('ProcessingTime', 5)
        packing_time = row['PackingTime']
        quantity = row['Quantity']
        sku = row['SKU']
        last_end = lane_last_end.get(lane, release)
        start_time = max(last_end, release)
        if 'Wave' in row:
            wave = row['Wave']
            prev_wave_min = orders_sorted[orders_sorted['Wave']==wave-1]['ReleaseTime'].min() if wave>1 else release
            start_time = min(start_time, current_time + timedelta(minutes=prev_wave_min)*theta)
        travel_time = compute_travel_time(quantity, lane_speed)
        induction_time = compute_induction_time(sku)
        completion_time = compute_completion_time(start_time, travel_time, processing_time, packing_time, induction_time)
        lane_last_end[lane] = completion_time
        lane_total_time[lane] = lane_total_time.get(lane, timedelta(0)) + (completion_time - start_time)
        sla_time = release + timedelta(hours=2)
        tardiness = max(timedelta(0), completion_time - sla_time)
        active_trays += 1
        if active_trays / len(orders_sorted) > Umax:
            print(f"Warning: potential gridlock at Order {row['OrderID']} (Ut>{Umax})")
        results.append({
            'OrderID': row['OrderID'],
            'Wave': row['Wave'],
            'Lane': lane,
            'StartTime': start_time,
            'CompletionTime': completion_time,
            'TravelTime': travel_time,
            'InductionTime': induction_time,
            'SLA': sla_time,
            'Tardiness': tardiness
        })
    lane_avg_time = sum(lane_total_time.values(), timedelta(0)) / len(lane_total_time)
    for r in results:
        lane_time = lane_total_time[r['Lane']]
        r['LaneImbalance'] = beta * abs(lane_time - lane_avg_time)
    return results

def run_pipeline_from_excel(file_path="orders.xlsx"):
    df = pd.read_excel(file_path)
    results = decode(df)
    return results

if __name__ == "__main__":
    input_file = "orders.xlsx"
    print(f"Reading orders from {input_file} ...")
    try:
        results = run_pipeline_from_excel(input_file)
        print("=== DLSSP Simulation Results ===")
        for r in results:
            print(
                f"Order {r['OrderID']} | Wave {r['Wave']} | Lane {r['Lane']} | "
                f"Start {r['StartTime']} | Completion {r['CompletionTime']} | "
                f"Travel {r['TravelTime']} | Induction {r['InductionTime']} | "
                f"Tardiness {r['Tardiness']} | LaneImbalance {r['LaneImbalance']}"
            )
        df_results = pd.DataFrame(results)
        output_file = "results_full.xlsx"
        df_results.to_excel(output_file, index=False)
        print(f"\nResults saved to {output_file}")
    except FileNotFoundError:
        print(f"ERROR: File '{input_file}' not found. Please create orders.xlsx first.")
