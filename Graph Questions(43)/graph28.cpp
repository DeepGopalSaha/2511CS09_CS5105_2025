#include <bits/stdc++.h>
using namespace std;
vector<int> shortestPathUnit(int n, vector<vector<int>>& edges, int src=0){
    vector<vector<int>> adj(n);
    for(auto &e: edges){ adj[e[0]].push_back(e[1]); adj[e[1]].push_back(e[0]); }
    vector<int> dist(n, -1); queue<int>q; dist[src]=0; q.push(src);
    while(!q.empty()){
        int u=q.front(); q.pop();
        for(int v: adj[u]) if(dist[v]==-1){ dist[v]=dist[u]+1; q.push(v); }
    }
    return dist;
}
int main(){ return 0; }
