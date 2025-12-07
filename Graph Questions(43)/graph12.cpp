#include <bits/stdc++.h>
using namespace std;
void dfs(int u, vector<int>& vis, vector<int>& order, vector<vector<int>>& adj){
    vis[u]=1; order.push_back(u+1);
    for(int v: adj[u]) if(!vis[v]) dfs(v,vis,order,adj);
}
vector<int> dfsTraversal(int V, vector<vector<int>>& adj){
    vector<int> vis(V,0), order;
    for(int i=0;i<V;i++) if(!vis[i]) dfs(i,vis,order,adj);
    return order;
}
int main(){ return 0; }
