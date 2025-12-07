#include <bits/stdc++.h>
using namespace std;
vector<int> topoDFS(int V, vector<vector<int>>& adj){
    vector<int> vis(V,0), order;
    function<void(int)> dfs = [&](int u){
        vis[u]=1;
        for(int v: adj[u]) if(!vis[v]) dfs(v);
        order.push_back(u);
    };
    for(int i=0;i<V;i++) if(!vis[i]) dfs(i);
    reverse(order.begin(), order.end());
    return order;
}
int main(){ return 0; }
