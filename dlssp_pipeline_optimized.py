import pandas as pd
from collections import defaultdict

class Order:
    def __init__(self, order_id, sku, wave, release_time, packing_time, lane=None):
        self.order_id = order_id
        self.sku = sku
        self.wave = wave
        self.release_time = release_time
        self.packing_time = packing_time
        self.lane = lane
        self.start_time = 0
        self.completion_time = 0
        self.travel_time = 0
        self.wait_time = 0
        self.sla_violation = False
        self.tardiness = 0

class Lane:
    def __init__(self, lane_id, speed=1.0):
        self.lane_id = lane_id
        self.speed = speed
        self.assigned_orders = []

def load_orders_from_excel(file_path):
    df = pd.read_excel(file_path)
    orders = []
    for _, row in df.iterrows():
        orders.append(Order(
            order_id=row['OrderID'],
            sku=row['SKU'],
            wave=row['Wave'],
            release_time=row['ReleaseTime'],
            packing_time=row['PackingTime']
        ))
    return orders

def assign_lanes(orders, lanes):
    lane_load = defaultdict(float)
    for order in sorted(orders, key=lambda x: x.release_time):
        lane = min(lanes, key=lambda l: lane_load[l.lane_id])
        order.lane = lane.lane_id
        lane.assigned_orders.append(order)
        lane_load[lane.lane_id] += order.packing_time
    return orders

def schedule_orders(orders, sorter_speed=1.0, max_utilization=0.85):
    waves = sorted(set(o.wave for o in orders))
    current_time = 0
    lane_end_time = defaultdict(float)

    for wave in waves:
        wave_orders = [o for o in orders if o.wave == wave]
        wave_start = min(o.release_time for o in wave_orders)
        for order in wave_orders:
            lane = order.lane
            start = max(wave_start, lane_end_time[lane])
            order.travel_time = 1.0 / sorter_speed
            order.start_time = start
            order.completion_time = start + order.travel_time + order.packing_time
            order.wait_time = start - order.release_time
            lane_end_time[lane] = order.completion_time
            order.sla_violation = order.completion_time > order.release_time + 10
            order.tardiness = max(0, order.completion_time - (order.release_time + 10))
            utilization = sum(o.packing_time for o in orders if o.start_time <= current_time < o.completion_time)
            if utilization > max_utilization:
                print(f"Warning: sorter utilization exceeded at time {current_time}")
    return orders

def run_pipeline_from_excel(file_path):
    lanes = [Lane(lane_id=i+1, speed=1.0) for i in range(3)]
    orders = load_orders_from_excel(file_path)
    orders = assign_lanes(orders, lanes)
    orders = schedule_orders(orders)
    results = []
    for o in orders:
        results.append({
            "OrderID": o.order_id,
            "Wave": o.wave,
            "Lane": o.lane,
            "Start": round(o.start_time, 2),
            "Completion": round(o.completion_time, 2),
            "Travel": round(o.travel_time, 2),
            "Wait": round(o.wait_time, 2),
            "SLA_violation": o.sla_violation,
            "Tardiness": round(o.tardiness, 2)
        })
    return results
