#include <bits/stdc++.h>
using namespace std;
bool canFinish(int N, vector<pair<int,int>>& prerequisites){
    vector<vector<int>> adj(N);
    vector<int> indeg(N,0);
    for(auto &p: prerequisites){ adj[p.second].push_back(p.first); indeg[p.first]++; }
    queue<int>q; for(int i=0;i<N;i++) if(indeg[i]==0) q.push(i);
    int cnt=0;
    while(!q.empty()){ int u=q.front(); q.pop(); cnt++; for(int v: adj[u]) if(--indeg[v]==0) q.push(v); }
    return cnt==N;
}
vector<int> findOrder(int N, vector<pair<int,int>>& prerequisites){
    vector<vector<int>> adj(N); vector<int> indeg(N,0);
    for(auto &p: prerequisites){ adj[p.second].push_back(p.first); indeg[p.first]++; }
    queue<int>q; for(int i=0;i<N;i++) if(indeg[i]==0) q.push(i);
    vector<int> order;
    while(!q.empty()){ int u=q.front(); q.pop(); order.push_back(u); for(int v: adj[u]) if(--indeg[v]==0) q.push(v); }
    if(order.size()!=N) return {}; return order;
}
int main(){ return 0; }
