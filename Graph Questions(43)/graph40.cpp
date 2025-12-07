#include <bits/stdc++.h>
using namespace std;
int largestIsland(vector<vector<int>>& grid){
    int n=grid.size();
    vector<int> area; area.push_back(0);
    vector<vector<int>> id(n, vector<int>(n,0));
    int idx=1;
    int dirs[4][2]={{1,0},{-1,0},{0,1},{0,-1}};
    for(int i=0;i<n;i++) for(int j=0;j<n;j++) if(grid[i][j]==1 && id[i][j]==0){
        int a=0; queue<pair<int,int>>q; q.push({i,j}); id[i][j]=idx;
        while(!q.empty()){
            auto [r,c]=q.front(); q.pop(); a++;
            for(auto &d:dirs){
                int nr=r+d[0], nc=c+d[1];
                if(nr<0||nr>=n||nc<0||nc>=n) continue;
                if(grid[nr][nc]==1 && id[nr][nc]==0){ id[nr][nc]=idx; q.push({nr,nc}); }
            }
        }
        area.push_back(a); idx++;
    }
    int ans=0;
    for(int i=1;i<idx;i++) ans=max(ans, area[i]);
    for(int i=0;i<n;i++) for(int j=0;j<n;j++) if(grid[i][j]==0){
        unordered_set<int> s; int sum=1;
        for(auto &d:dirs){
            int nr=i+d[0], nc=j+d[1];
            if(nr<0||nr>=n||nc<0||nc>=n) continue;
            if(id[nr][nc]) s.insert(id[nr][nc]);
        }
        for(int idc: s) sum += area[idc];
        ans = max(ans, sum);
    }
    return ans;
}
int main(){ return 0; }
