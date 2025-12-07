#include <bits/stdc++.h>
using namespace std;
bool dfsCycle(int u, int p, vector<int>& vis, vector<vector<int>>& adj){
    vis[u]=1;
    for(int v: adj[u]){
        if(!vis[v]){ if(dfsCycle(v,u,vis,adj)) return true; }
        else if(v!=p) return true;
    }
    return false;
}
bool isCycleDFS(int V, vector<vector<int>>& adj){
    vector<int> vis(V,0);
    for(int i=0;i<V;i++) if(!vis[i]) if(dfsCycle(i,-1,vis,adj)) return true;
    return false;
}
int main(){ return 0; }
