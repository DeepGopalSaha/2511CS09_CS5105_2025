#include <bits/stdc++.h>
using namespace std;
vector<int> eventualSafeNodes(int V, vector<vector<int>>& adj){
    vector<vector<int>> rev(V);
    vector<int> outdeg(V,0);
    for(int u=0;u<V;u++){
        outdeg[u]=adj[u].size();
        for(int v: adj[u]) rev[v].push_back(u);
    }
    queue<int>q; vector<int> safe(V,0);
    for(int i=0;i<V;i++) if(outdeg[i]==0) { q.push(i); safe[i]=1; }
    while(!q.empty()){
        int u=q.front(); q.pop();
        for(int v: rev[u]){
            if(--outdeg[v]==0){ safe[v]=1; q.push(v); }
        }
    }
    vector<int> res;
    for(int i=0;i<V;i++) if(safe[i]) res.push_back(i);
    sort(res.begin(), res.end()); return res;
}
int main(){ return 0; }
