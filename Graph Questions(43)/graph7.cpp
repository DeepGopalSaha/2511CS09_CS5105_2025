#include <bits/stdc++.h>
using namespace std;
int shortestPathBinaryMatrix(vector<vector<int>>& grid, pair<int,int> src, pair<int,int> dst){
    int n=grid.size(), m=grid[0].size();
    if(grid[src.first][src.second]==0 || grid[dst.first][dst.second]==0) return -1;
    vector<vector<int>> dist(n, vector<int>(m, -1));
    queue<pair<int,int>>q; q.push(src); dist[src.first][src.second]=0;
    int dirs[4][2]={{1,0},{-1,0},{0,1},{0,-1}};
    while(!q.empty()){
        auto [r,c]=q.front(); q.pop();
        if(r==dst.first && c==dst.second) return dist[r][c];
        for(auto &d:dirs){
            int nr=r+d[0], nc=c+d[1];
            if(nr<0||nr>=n||nc<0||nc>=m) continue;
            if(grid[nr][nc]==1 && dist[nr][nc]==-1){ dist[nr][nc]=dist[r][c]+1; q.push({nr,nc}); }
        }
    }
    return -1;
}
int main(){ return 0; }
