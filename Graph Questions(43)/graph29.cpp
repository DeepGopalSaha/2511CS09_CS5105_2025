#include <bits/stdc++.h>
using namespace std;
vector<long long> shortestPathDAG(int n, vector<vector<array<int,3>>>& edges, int src=0){
    vector<vector<pair<int,int>>> adj(n);
    for(auto &e: edges) adj[e[0]].push_back({e[1], e[2]});
    vector<int> indeg(n,0);
    for(int u=0;u<n;u++) for(auto &p: adj[u]) indeg[p.first]++;
    queue<int>q; for(int i=0;i<n;i++) if(indeg[i]==0) q.push(i);
    vector<int> topo;
    while(!q.empty()){ int u=q.front(); q.pop(); topo.push_back(u); for(auto &p: adj[u]) if(--indeg[p.first]==0) q.push(p.first); }
    const long long INF = 9e18;
    vector<long long> dist(n, INF); dist[src]=0;
    for(int u: topo){
        if(dist[u]==INF) continue;
        for(auto &p: adj[u]){
            int v=p.first, w=p.second;
            if(dist[v] > dist[u] + w) dist[v] = dist[u] + w;
        }
    }
    for(int i=0;i<n;i++) if(dist[i]==INF) dist[i] = -1;
    return dist;
}
int main(){ return 0; }
