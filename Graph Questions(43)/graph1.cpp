#include <bits/stdc++.h>
using namespace std;
vector<int> bfsTraversal(int V, vector<vector<int>>& adj) {
    vector<int> vis(V,0), order;
    for(int start=0; start<V; ++start){
        if(vis[start]) continue;
        queue<int>q; q.push(start); vis[start]=1;
        while(!q.empty()){
            int u=q.front(); q.pop();
            order.push_back(u+1);
            for(int v: adj[u]) if(!vis[v]) { vis[v]=1; q.push(v); }
        }
    }
    return order;
}
int main(){ ios::sync_with_stdio(false); cin.tie(nullptr);
    return 0;
}
