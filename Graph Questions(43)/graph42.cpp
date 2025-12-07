#include <bits/stdc++.h>
using namespace std;
vector<int> articulationPoints(int n, vector<vector<int>>& adj){
    vector<int> tin(n,-1), low(n,-1), vis(n,0), isArt(n,0);
    int timer=0;
    function<void(int,int)> dfs = [&](int v, int p){
        vis[v]=1; tin[v]=low[v]=timer++; int children=0;
        for(int to: adj[v]){
            if(to==p) continue;
            if(vis[to]) low[v]=min(low[v], tin[to]);
            else{
                dfs(to,v); children++;
                low[v]=min(low[v], low[to]);
                if(p!=-1 && low[to] >= tin[v]) isArt[v]=1;
            }
        }
        if(p==-1 && children>1) isArt[v]=1;
    };
    for(int i=0;i<n;i++) if(!vis[i]) dfs(i,-1);
    vector<int> res;
    for(int i=0;i<n;i++) if(isArt[i]) res.push_back(i);
    return res;
}
int main(){ return 0; }
