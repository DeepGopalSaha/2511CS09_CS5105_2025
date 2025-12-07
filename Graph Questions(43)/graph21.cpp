#include <bits/stdc++.h>
using namespace std;
bool dfs(int u, vector<int>& vis, vector<vector<int>>& adj){
    vis[u]=1;
    for(int v: adj[u]){
        if(vis[v]==0){ if(dfs(v,vis,adj)) return true; }
        else if(vis[v]==1) return true;
    }
    vis[u]=2; return false;
}
bool isCyclic(int V, vector<vector<int>>& adj){
    vector<int> vis(V,0);
    for(int i=0;i<V;i++) if(vis[i]==0) if(dfs(i,vis,adj)) return true;
    return false;
}
int main(){ return 0; }
