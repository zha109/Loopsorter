import math
import pandas as pd
import networkx as nx
from collections import defaultdict
from datetime import datetime

def build_order_graph(M: pd.DataFrame, weight_mode: str = "jaccard", min_weight: float = 0.0) -> nx.Graph:
    orders = list(M.index)
    G = nx.Graph()
    G.add_nodes_from(orders)
    bin_sets = {o: set(M.columns[M.loc[o] > 0]) for o in orders}
    for i in range(len(orders)):
        for j in range(i+1, len(orders)):
            o1, o2 = orders[i], orders[j]
            inter = bin_sets[o1] & bin_sets[o2]
            if weight_mode == "shared":
                w = float(len(inter))
            elif weight_mode == "jaccard":
                uni = bin_sets[o1] | bin_sets[o2]
                w = float(len(inter))/float(len(uni)) if len(uni) > 0 else 0.0
            else:
                raise ValueError("weight_mode must be 'shared' or 'jaccard'.")
            if w > min_weight:
                G.add_edge(o1,o2,weight=w)
    return G

def partition_to_groups(part: dict) -> dict:
    groups = defaultdict(list)
    for node, lab in part.items():
        groups[lab].append(node)
    for k in groups:
        groups[k].sort()
    return dict(sorted(groups.items(), key=lambda kv: kv[0]))

def _weighted_degree(G, n):
    return sum(d.get('weight',1.0) for _, _, d in G.edges(n, data=True))

def _modularity_gamma(G, part, gamma=1.0):
    m = sum(d.get('weight',1.0) for *_, d in G.edges(data=True))
    if m == 0: return 0.0
    tot = defaultdict(float)
    intra = defaultdict(float)
    for n,c in part.items():
        tot[c] += _weighted_degree(G,n)
    for u,v,d in G.edges(data=True):
        if part[u]==part[v]:
            intra[part[u]] += d.get('weight',1.0)
    Q = 0.0
    for c in set(part.values()):
        Sigma_in = intra.get(c,0.0)
        Sigma_tot = tot.get(c,0.0)
        Q += (Sigma_in/(2*m)) - gamma*(Sigma_tot**2)/(4*m*m)
    return Q

def _community_bin_union(M: pd.DataFrame, members: list) -> set:
    if not members: return set()
    sub = M.loc[members]
    return set(sub.columns[(sub.sum(axis=0) > 0)])

def _jaccard(a: set, b: set) -> float:
    if not a and not b: return 1.0
    if not a or not b: return 0.0
    return len(a & b)/len(a | b)

def _violations_of(part,max_size:int)->int:
    sizes=defaultdict(int)
    for _,cid in part.items(): sizes[cid]+=1
    return sum(max(0,s-max_size) for s in sizes.values())

def louvain_phase1_verbose(G,M,gamma=1.0,eps=1e-12):
    def _tie_tuple_P1(u, cand_cid, part):
        members=[x for x,c in part.items() if c==cand_cid]
        k_in=sum(G[u][v]['weight'] for v in members if G.has_edge(u,v))
        deg_u=sum(d.get('weight',1.0) for _,_,d in G.edges(u,data=True))
        k_out=max(deg_u-k_in,0.0)
        Bu=set(M.columns[M.loc[u]>0])
        Uc=_community_bin_union(M,members)
        jac=_jaccard(Bu,Uc)
        return (k_in,jac,-k_out,-cand_cid)
    
    nodes=list(G.nodes())
    part={n:i for i,n in enumerate(nodes)}
    while True:
        moved_this_pass=0
        for u in nodes:
            cu=part[u]
            neighbor_cids=sorted(set(part[v] for v in G.neighbors(u))-{cu})
            if not neighbor_cids: continue
            baseQ=_modularity_gamma(G,part,gamma)
            dqs={cand:_modularity_gamma(G,{**part,u:cand},gamma)-baseQ for cand in neighbor_cids}
            best_dQ=max(dqs.values())
            if best_dQ<=eps: continue
            best_cands=[c for c,v in dqs.items() if abs(v-best_dQ)<=eps]
            chosen=best_cands[0] if len(best_cands)==1 else max(best_cands,key=lambda c:_tie_tuple_P1(u,c,part))
            part[u]=chosen
            moved_this_pass+=1
        if moved_this_pass==0: break
    return part

def _aggregate_graph(G, part):
    comm_nodes=defaultdict(list)
    for n,c in part.items(): comm_nodes[c].append(n)
    H=nx.Graph()
    for c in comm_nodes: H.add_node(c)
    for c,members in comm_nodes.items():
        w_self=0.0
        for i in range(len(members)):
            for j in range(i+1,len(members)):
                u,v=members[i],members[j]
                if G.has_edge(u,v): w_self+=G[u][v]['weight']
        if w_self>0: H.add_edge(c,c,weight=w_self)
    comm_ids=sorted(comm_nodes.keys())
    for i in range(len(comm_ids)):
        for j in range(i+1,len(comm_ids)):
            c1,c2=comm_ids[i],comm_ids[j]
            w=0.0
            for u in comm_nodes[c1]:
                for v in comm_nodes[c2]:
                    if G.has_edge(u,v): w+=G[u][v]['weight']
            if w>0: H.add_edge(c1,c2,weight=w)
    return H, comm_nodes

def louvain_phase2_verbose(G, part_after_p1, gamma=1.0, eps=1e-12):
    H, comm_nodes=_aggregate_graph(G,part_after_p1)
    nodesH=list(H.nodes())
    partH={n:i for i,n in enumerate(nodesH)}
    def mod_H(H,partH,gamma=1.0):
        mH=H.size(weight='weight')
        if mH==0: return 0.0
        degH=dict(H.degree(weight='weight'))
        tot=defaultdict(float)
        intra=defaultdict(float)
        for n,c in partH.items(): tot[c]+=degH[n]
        for u,v,d in H.edges(data=True):
            if partH[u]==partH[v]: intra[partH[u]]+=d.get('weight',1.0)
        Q=0.0
        for c in set(partH.values()):
            Q+=(intra.get(c,0.0)/(2*mH))-gamma*(tot.get(c,0.0)**2)/(4*mH*mH)
        return Q
    while True:
        moved_this_pass=0
        for u in nodesH:
            cu=partH[u]
            neighbor_cids=sorted(set(partH[v] for v in H.neighbors(u))-{cu})
            if not neighbor_cids: continue
            baseQ=mod_H(H,partH,gamma)
            dqs={cand:mod_H(H,{**partH,u:cand},gamma)-baseQ for cand in neighbor_cids}
            best_dQ=max(dqs.values())
            if best_dQ<=eps: continue
            best_cands=[c for c,v in dqs.items() if abs(v-best_dQ)<=eps]
            chosen=best_cands[0] if len(best_cands)==1 else min(best_cands)
            partH[u]=chosen
            moved_this_pass+=1
        if moved_this_pass==0: break
    comm_of_H=defaultdict(list)
    for nH,cH in partH.items():
        for orig in comm_nodes[nH]: comm_of_H[cH].append(orig)
    part_back={}
    for cid,members in comm_of_H.items():
        for n in members: part_back[n]=cid
    return part_back

def improve_with_lexi_tiebreak(G,M,part,eps_mod=1e-12,max_iters=5,target_size=None,min_size=None,max_size=None,
                               penalty_lambda=1e6,allow_new_community=True,gamma=1.0):
    cur_part=dict(part)
    def _obj(p): return _modularity_gamma(G,p,gamma)-penalty_lambda*_violations_of(p,max_size) if max_size else _modularity_gamma(G,p,gamma)
    for it in range(1,max_iters+1):
        moved_this_round=0
        next_cid=(max(cur_part.values())+1) if cur_part else 0
        for u in list(cur_part.keys()):
            cid_u=cur_part[u]
            neighbor_cids=set(cur_part[v] for v in G.neighbors(u))
            candidate_cids=sorted(neighbor_cids|{cid_u})
            if allow_new_community: candidate_cids.append(next_cid)
            feasible=[cid for cid in candidate_cids if (cid==cid_u) or (max_size is None or list(cur_part.values()).count(cid)<max_size)]
            if not feasible: continue
            objs={cid:_obj({**cur_part,u:cid}) for cid in feasible}
            best_obj=max(objs.values())
            best_cids=[cid for cid,v in objs.items() if abs(v-best_obj)<=eps_mod]
            chosen=best_cids[0]
            if chosen!=cid_u: cur_part[u]=chosen; moved_this_round+=1
        if moved_this_round==0: break
    final_mod=_modularity_gamma(G,cur_part,gamma)
    return cur_part,final_mod,moved_this_round

def run_pipeline_from_excel(file_path):
    M=pd.read_excel(file_path,sheet_name="Incidence",index_col=0,engine="openpyxl")
    G=build_order_graph(M,weight_mode="jaccard",min_weight=0.0)
    print(f"[INFO] Graph -> nodes: {G.number_of_nodes()} edges: {G.number_of_edges()}")
    p1=louvain_phase1_verbose(G,M)
    p2=louvain_phase2_verbose(G,p1)
    part_final,mod_final,nmove=improve_with_lexi_tiebreak(G,M,p2,max_iters=5)
    print(f"[INFO] Final modularity Q={mod_final:.6f} | refinement moves={nmove}")
    print("[INFO] Partition final:", partition_to_groups(part_final))

def run_dlssp_example():
    df_incidence=pd.DataFrame({"sku1":[1,0,1],"sku2":[0,1,1],"sku3":[1,1,0]},index=["order1","order2","order3"])
    df_orders=pd.DataFrame({"priority":["Express","Fast","Standard"],"due_date":[pd.Timestamp("2025-12-05"),pd.Timestamp("2025-12-06"),pd.Timestamp("2025-12-07")]},index=["order1","order2","order3"])
    path="dlssp_example.xlsx"
    with pd.ExcelWriter(path,engine="xlsxwriter") as w:
        df_incidence.to_excel(w,sheet_name="Incidence")
        df_orders.to_excel(w,sheet_name="Orders")
    run_pipeline_from_excel(path)
    return path

if __name__=="__main__":
    run_dlssp_example()
