#include <bits/stdc++.h>
using namespace std;
bool dfsColor(int u, vector<int>& color, vector<vector<int>>& adj){
    for(int v: adj[u]){
        if(color[v]==-1){
            color[v]=1-color[u];
            if(!dfsColor(v,color,adj)) return false;
        } else if(color[v]==color[u]) return false;
    }
    return true;
}
bool isBipartite(int V, vector<vector<int>>& adj){
    vector<int> color(V, -1);
    for(int i=0;i<V;i++) if(color[i]==-1){
        color[i]=0;
        if(!dfsColor(i,color,adj)) return false;
    }
    return true;
}
int main(){ return 0; }
