#include <bits/stdc++.h>
using namespace std;
int numIslands(vector<vector<char>>& grid){
    int n=grid.size(); if(n==0) return 0; int m=grid[0].size();
    int cnt=0;
    int dirs[8][2]={{1,0},{-1,0},{0,1},{0,-1},{1,1},{1,-1},{-1,1},{-1,-1}};
    for(int i=0;i<n;i++) for(int j=0;j<m;j++) if(grid[i][j]=='1'){
        cnt++; queue<pair<int,int>>q; q.push({i,j}); grid[i][j]='0';
        while(!q.empty()){
            auto [r,c]=q.front(); q.pop();
            for(auto &d:dirs){
                int nr=r+d[0], nc=c+d[1];
                if(nr<0||nr>=n||nc<0||nc>=m) continue;
                if(grid[nr][nc]=='1'){ grid[nr][nc]='0'; q.push({nr,nc}); }
            }
        }
    }
    return cnt;
}
int main(){ return 0; }
