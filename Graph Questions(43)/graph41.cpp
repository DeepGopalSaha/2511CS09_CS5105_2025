#include <bits/stdc++.h>
using namespace std;
vector<pair<int,int>> bridgesTarjan(int n, vector<vector<int>>& adj){
    vector<int> tin(n,-1), low(n,-1), vis(n,0);
    vector<pair<int,int>> bridges;
    int timer=0;
    function<void(int,int)> dfs = [&](int v, int p){
        vis[v]=1; tin[v]=low[v]=timer++;
        for(int to: adj[v]){
            if(to==p) continue;
            if(vis[to]) low[v] = min(low[v], tin[to]);
            else{
                dfs(to, v);
                low[v] = min(low[v], low[to]);
                if(low[to] > tin[v]) bridges.push_back({v,to});
            }
        }
    };
    for(int i=0;i<n;i++) if(!vis[i]) dfs(i,-1);
    return bridges;
}
int main(){ return 0; }
