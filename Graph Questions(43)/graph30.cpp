#include <bits/stdc++.h>
using namespace std;
vector<long long> dijkstraSet(int V, vector<vector<pair<int,int>>>& adj, int S){
    const long long INF = 9e18;
    vector<long long> dist(V, INF);
    set<pair<long long,int>> st;
    dist[S]=0; st.insert({0,S});
    while(!st.empty()){
        auto it = st.begin(); auto [d,u]=*it; st.erase(it);
        for(auto &e: adj[u]){
            int v=e.first, w=e.second;
            if(dist[v] > d + w){
                if(dist[v]!=INF) st.erase({dist[v], v});
                dist[v] = d + w; st.insert({dist[v], v});
            }
        }
    }
    return dist;
}
int main(){ return 0; }
