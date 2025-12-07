#include <bits/stdc++.h>
using namespace std;
vector<int> kahnTopo(int V, vector<vector<int>>& adj){
    vector<int> indeg(V,0), order;
    for(int u=0;u<V;u++) for(int v: adj[u]) indeg[v]++;
    queue<int>q; for(int i=0;i<V;i++) if(indeg[i]==0) q.push(i);
    while(!q.empty()){
        int u=q.front(); q.pop(); order.push_back(u);
        for(int v: adj[u]) if(--indeg[v]==0) q.push(v);
    }
    if(order.size()!=V) return {};
    return order;
}
int main(){ return 0; }
