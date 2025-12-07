#include <bits/stdc++.h>
using namespace std;
bool isCycle(int V, vector<vector<int>>& adj){
    vector<int> vis(V,0);
    for(int i=0;i<V;i++) if(!vis[i]){
        queue<pair<int,int>>q; q.push({i,-1}); vis[i]=1;
        while(!q.empty()){
            auto [u,par]=q.front(); q.pop();
            for(int v: adj[u]){
                if(!vis[v]){ vis[v]=1; q.push({v,u}); }
                else if(v!=par) return true;
            }
        }
    }
    return false;
}
int main(){ return 0; }
