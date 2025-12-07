#include <bits/stdc++.h>
using namespace std;
const int MOD = 1e9+7;
int countPaths(int n, vector<vector<array<int,3>>>& adj){
    const long long INF = 9e18;
    vector<long long> dist(n, INF), ways(n,0);
    priority_queue<pair<long long,int>, vector<pair<long long,int>>, greater<pair<long long,int>>> pq;
    dist[0]=0; ways[0]=1; pq.push({0,0});
    while(!pq.empty()){
        auto [d,u]=pq.top(); pq.pop();
        if(d>dist[u]) continue;
        for(auto &e: adj[u]){
            int v=e[0], w=e[1];
            if(dist[v] > d + w){
                dist[v] = d + w;
                ways[v] = ways[u];
                pq.push({dist[v], v});
            } else if(dist[v] == d + w){
                ways[v] = (ways[v] + ways[u]) % MOD;
            }
        }
    }
    return (int)ways[n-1];
}
int main(){ return 0; }
