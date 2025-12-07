#include <bits/stdc++.h>
using namespace std;
vector<long long> bellmanFord(int V, vector<array<int,3>>& edges, int S){
    const long long INF = 9e18;
    vector<long long> dist(V, INF); dist[S]=0;
    for(int i=0;i<V-1;i++){
        for(auto &e: edges){
            int u=e[0], v=e[1], w=e[2];
            if(dist[u]!=INF && dist[v] > dist[u] + w) dist[v] = dist[u] + w;
        }
    }
    for(auto &e: edges){
        int u=e[0], v=e[1], w=e[2];
        if(dist[u]!=INF && dist[v] > dist[u] + w) return vector<long long>{-1};
    }
    return dist;
}
int main(){ return 0; }
