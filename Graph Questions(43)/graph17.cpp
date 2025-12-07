#include <bits/stdc++.h>
using namespace std;
int numEnclaves(vector<vector<int>>& grid){
    int n=grid.size(), m=grid[0].size();
    queue<pair<int,int>>q;
    for(int i=0;i<n;i++) for(int j=0;j<m;j++){
        if((i==0||j==0||i==n-1||j==m-1) && grid[i][j]==1){ q.push({i,j}); grid[i][j]=0; }
    }
    int dirs[4][2]={{1,0},{-1,0},{0,1},{0,-1}};
    while(!q.empty()){
        auto [r,c]=q.front(); q.pop();
        for(auto &d:dirs){
            int nr=r+d[0], nc=c+d[1];
            if(nr<0||nr>=n||nc<0||nc>=m) continue;
            if(grid[nr][nc]==1){ grid[nr][nc]=0; q.push({nr,nc}); }
        }
    }
    int cnt=0; for(int i=0;i<n;i++) for(int j=0;j<m;j++) if(grid[i][j]==1) cnt++;
    return cnt;
}
int main(){ return 0; }
