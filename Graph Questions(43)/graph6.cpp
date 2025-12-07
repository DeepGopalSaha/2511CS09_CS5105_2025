#include <bits/stdc++.h>
using namespace std;
vector<long long> dijkstra(int V, vector<vector<pair<int,int>>>& adj, int S){
    const long long INF = 9e18;
    vector<long long> dist(V, INF);
    priority_queue<pair<long long,int>, vector<pair<long long,int>>, greater<pair<long long,int>>> pq;
    dist[S]=0; pq.push({0,S});
    while(!pq.empty()){
        auto [d,u]=pq.top(); pq.pop();
        if(d>dist[u]) continue;
        for(auto &e: adj[u]){
            int v=e.first; int w=e.second;
            if(dist[v]>d+w){ dist[v]=d+w; pq.push({dist[v], v}); }
        }
    }
    return dist;
}
int main(){ return 0; }
