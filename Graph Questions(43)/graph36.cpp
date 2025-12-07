#include <bits/stdc++.h>
using namespace std;
long long primMST(int V, vector<vector<pair<int,int>>>& adj){
    vector<int> vis(V,0);
    priority_queue<pair<int,int>, vector<pair<int,int>>, greater<pair<int,int>>> pq;
    pq.push({0,0});
    long long ans=0;
    while(!pq.empty()){
        auto [w,u]=pq.top(); pq.pop();
        if(vis[u]) continue;
        vis[u]=1; ans += w;
        for(auto &e: adj[u]) if(!vis[e.first]) pq.push({e.second, e.first});
    }
    return ans;
}
int main(){ return 0; }
