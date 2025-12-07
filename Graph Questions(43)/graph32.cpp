#include <bits/stdc++.h>
using namespace std;
int minMultiplications(vector<int>& arr, int start, int end){
    const int MOD = 100000;
    vector<int> dist(MOD, -1);
    queue<int>q; dist[start]=0; q.push(start);
    while(!q.empty()){
        int u=q.front(); q.pop();
        if(u==end) return dist[u];
        for(int x: arr){
            int v = (u * 1LL * x) % MOD;
            if(dist[v]==-1){ dist[v]=dist[u]+1; q.push(v); }
        }
    }
    return -1;
}
int main(){ return 0; }
