#include <bits/stdc++.h>
using namespace std;
int findProvinces(int V, vector<vector<int>>& adj) {
    vector<int> vis(V,0); int cnt=0;
    for(int i=0;i<V;i++){
        if(vis[i]) continue;
        cnt++;
        queue<int>q; q.push(i); vis[i]=1;
        while(!q.empty()){
            int u=q.front(); q.pop();
            for(int v: adj[u]) if(!vis[v]) { vis[v]=1; q.push(v); }
        }
    }
    return cnt;
}
int main(){ return 0; }
