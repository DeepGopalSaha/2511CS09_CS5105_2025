#include <bits/stdc++.h>
using namespace std;
int minimumEffortPath(vector<vector<int>>& heights){
    int n=heights.size(), m=heights[0].size();
    const int INF=1e9;
    vector<vector<int>> dist(n, vector<int>(m, INF));
    priority_queue< tuple<int,int,int>, vector<tuple<int,int,int>>, greater<tuple<int,int,int>> > pq;
    dist[0][0]=0; pq.push({0,0,0});
    int dirs[4][2]={{1,0},{-1,0},{0,1},{0,-1}};
    while(!pq.empty()){
        auto [cost,r,c]=pq.top(); pq.pop();
        if(cost>dist[r][c]) continue;
        if(r==n-1 && c==m-1) return cost;
        for(auto &d:dirs){
            int nr=r+d[0], nc=c+d[1];
            if(nr<0||nr>=n||nc<0||nc>=m) continue;
            int nd = max(cost, abs(heights[r][c]-heights[nr][nc]));
            if(nd < dist[nr][nc]){ dist[nr][nc]=nd; pq.push({nd,nr,nc}); }
        }
    }
    return 0;
}
int main(){ return 0; }
