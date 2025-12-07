#include <bits/stdc++.h>
using namespace std;
vector<vector<int>> kosaraju(int V, vector<vector<int>>& adj){
    vector<int> vis(V,0);
    vector<int> order;
    function<void(int)> dfs1 = [&](int u){ vis[u]=1; for(int v: adj[u]) if(!vis[v]) dfs1(v); order.push_back(u); };
    for(int i=0;i<V;i++) if(!vis[i]) dfs1(i);
    vector<vector<int>> rev(V);
    for(int u=0;u<V;u++) for(int v: adj[u]) rev[v].push_back(u);
    fill(vis.begin(), vis.end(), 0);
    vector<vector<int>> comps;
    function<void(int, vector<int>&)> dfs2 = [&](int u, vector<int>& comp){ vis[u]=1; comp.push_back(u); for(int v: rev[u]) if(!vis[v]) dfs2(v, comp); };
    for(int i=V-1;i>=0;i--){
        int v = order[i];
        if(!vis[v]){ vector<int> comp; dfs2(v, comp); comps.push_back(comp); }
    }
    return comps;
}
int main(){ return 0; }
