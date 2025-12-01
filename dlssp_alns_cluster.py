import pandas as pd
from datetime import datetime, timedelta
import random
import math

def load_params(file_path="params.xlsx"):
    df = pd.read_excel(file_path)
    params = {}
    for _, row in df.iterrows():
        key = str(row['param']).strip()
        value = row['value']
        try:
            value = float(value)
        except:
            pass
        params[key] = value
    return params

def compute_travel_time(quantity, lane_speed, lane_pos, sku_pos, distance_factor=1.0):
    distance = abs(lane_pos - sku_pos) * distance_factor
    return timedelta(minutes=quantity * distance / lane_speed)

def compute_induction_time(sku):
    return timedelta(minutes=1 + (hash(sku) % 3))

def compute_completion_time(start, travel, processing, packing, induction):
    return start + travel + timedelta(minutes=processing) + timedelta(minutes=packing) + induction

def assign_tray(sku, lane, lane_positions):
    sku_pos = hash(sku) % 30
    lane_pos = lane_positions.get(lane, 0)
    return sku_pos, lane_pos

def schedule_orders(orders_df, params, lane_positions):
    results = []
    lane_last_end = {}
    lane_total_time = {}
    current_time = datetime.now().replace(hour=8, minute=0, second=0, microsecond=0)
    orders_sorted = orders_df.sort_values(by='ReleaseTime')
    active_trays = 0
    Umax = float(params.get('Umax', 0.85))
    theta = float(params.get('theta', 0.3))

    for _, row in orders_sorted.iterrows():
        lane = row['Lane']
        release = current_time + timedelta(minutes=row['ReleaseTime'])
        lane_speed = row['LaneSpeed']
        processing_time = row.get('ProcessingTime', 5)
        packing_time = row.get('PackingTime', 5)
        quantity = row['Quantity']
        sku = row['SKU']

        last_end = lane_last_end.get(lane, release)
        start_time = max(last_end, release)

        wave = row.get('Wave', 1)
        if wave > 1:
            prev_wave_min = orders_sorted[orders_sorted['Wave']==wave-1]['ReleaseTime'].min()
            if pd.notna(prev_wave_min):
                overlap_minutes = float(prev_wave_min) * theta
                start_time = min(start_time, current_time + timedelta(minutes=overlap_minutes))

        sku_pos, lane_pos = assign_tray(sku, lane, lane_positions)
        travel_time = compute_travel_time(quantity, lane_speed, lane_pos, sku_pos)
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
            'Wave': wave,
            'Lane': lane,
            'StartTime': start_time,
            'CompletionTime': completion_time,
            'TravelTime': travel_time,
            'InductionTime': induction_time,
            'TrayPos': sku_pos,
            'SLA': sla_time,
            'Tardiness': tardiness
        })

    lane_avg_time = sum(lt.total_seconds() for lt in lane_total_time.values()) / len(lane_total_time)
    for r in results:
        lane_time = lane_total_time[r['Lane']].total_seconds()
        beta = float(params.get('beta_l', 0.5))
        r['LaneImbalance'] = beta * abs(lane_time - lane_avg_time)/60

    return results

def alns_optimize(orders_df, params, lane_positions):
    best_solution = schedule_orders(orders_df, params, lane_positions)
    best_score = compute_objective(best_solution, params)

    alns_iters = int(params.get('alns_iters', 200))
    destroy_min = int(params.get('alns_destroy_k_min', 2))
    destroy_max = int(params.get('alns_destroy_k_max', 4))

    for _ in range(alns_iters):
        temp_df = orders_df.copy()
        remove_n = random.randint(destroy_min, destroy_max)
        remove_idx = random.sample(list(temp_df.index), min(remove_n, len(temp_df)))
        temp_df = temp_df.drop(remove_idx)

        temp_df = temp_df.sample(frac=1)
        temp_solution = schedule_orders(temp_df, params, lane_positions)
        score = compute_objective(temp_solution, params)
        if score < best_score:
            best_solution = temp_solution
            best_score = score

    return best_solution

def compute_objective(results, params):
    lambda1 = float(params.get('lambda1', 1e6))
    lambda2 = float(params.get('lambda2', 1000))
    lambda3 = float(params.get('lambda3', 1))
    Cmax = max(r['CompletionTime'] for r in results)
    lane_imbalance = sum(r['LaneImbalance'] for r in results)
    sla_penalty = sum(max(timedelta(0), r['CompletionTime'] - r['SLA']).total_seconds()/60 for r in results)
    tardiness_penalty = sum(r['Tardiness'].total_seconds()/60 for r in results)
    total_score = lambda3*Cmax.timestamp() + lambda2*lane_imbalance + lambda1*sla_penalty + tardiness_penalty
    return total_score

def run_pipeline(orders_file="orders.xlsx", params_file="params.xlsx"):
    orders_df = pd.read_excel(orders_file)
    params = load_params(params_file)
    lane_positions = {1:0, 2:10, 3:20}
    results = alns_optimize(orders_df, params, lane_positions)
    return results

if __name__ == "__main__":
    orders_file = "orders.xlsx"
    params_file = "params.xlsx"
    print(f"Reading orders from {orders_file} ...")
    try:
        results = run_pipeline(orders_file, params_file)
        print("=== DLSSP ALNS + Tray Clustering Simulation Results ===")
        for r in results:
            print(f"Order {r['OrderID']} | Wave {r['Wave']} | Lane {r['Lane']} | "
                  f"Start {r['StartTime']} | Completion {r['CompletionTime']} | "
                  f"Travel {r['TravelTime']} | Induction {r['InductionTime']} | "
                  f"Tardiness {r['Tardiness']} | LaneImbalance {r['LaneImbalance']:.2f}")
        df_results = pd.DataFrame(results)
        df_results.to_excel("results_alns_cluster.xlsx", index=False)
        print(f"\nResults saved to results_alns_cluster.xlsx")
    except FileNotFoundError as e:
        print(f"ERROR: {e}")
