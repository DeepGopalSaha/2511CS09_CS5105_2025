#include <bits/stdc++.h>
using namespace std;
int orangesRotting(vector<vector<int>>& grid) {
    int n=grid.size(), m=grid[0].size();
    queue<pair<int,int>>q;
    int fresh=0;
    for(int i=0;i<n;i++) for(int j=0;j<m;j++){
        if(grid[i][j]==2) q.push({i,j});
        else if(grid[i][j]==1) fresh++;
    }
    int minutes=0;
    int dirs[4][2]={{1,0},{-1,0},{0,1},{0,-1}};
    while(!q.empty() && fresh>0){
        int sz=q.size();
        for(int k=0;k<sz;k++){
            auto [r,c]=q.front(); q.pop();
            for(auto &d:dirs){
                int nr=r+d[0], nc=c+d[1];
                if(nr<0||nr>=n||nc<0||nc>=m) continue;
                if(grid[nr][nc]==1){ grid[nr][nc]=2; fresh--; q.push({nr,nc}); }
            }
        }
        minutes++;
    }
    return fresh? -1: minutes;
}
int main(){ return 0; }
