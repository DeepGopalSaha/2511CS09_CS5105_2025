#include <bits/stdc++.h>
using namespace std;
int cheapestFlights(int n, vector<vector<int>>& flights, int src, int dst, int K){
    vector<vector<pair<int,int>>> adj(n);
    for(auto &f: flights) adj[f[0]].push_back({f[1], f[2]});
    const long long INF = 9e18;
    vector<long long> dist(n, INF);
    queue<tuple<int,int,int>> q;
    q.push({0, src, 0});
    dist[src]=0;
    int answer = -1;
    vector<long long> best(n, INF);
    best[src]=0;
    while(!q.empty()){
        auto [cost,u,stops]=q.front(); q.pop();
        if(stops > K) continue;
        if(u==dst){ if(answer==-1 || cost<answer) answer = cost; }
        for(auto &e: adj[u]){
            int v=e.first, w=e.second;
            int nc = cost + w;
            if(nc < best[v] && nc < (answer==-1?INF:answer)){
                best[v]=nc;
                q.push({nc, v, stops+1});
            }
        }
    }
    return answer;
}
int main(){ return 0; }
