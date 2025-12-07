#include <bits/stdc++.h>
using namespace std;
vector<int> brute_spiral_order(vector<vector<int>>& a){
    int m=a.size(), n=a[0].size();
    vector<int> res; vector<vector<int>> used(m, vector<int>(n));
    int dir=0, r=0,c=0;
    int dr[4]={0,1,0,-1}, dc[4]={1,0,-1,0};
    for(int k=0;k<m*n;++k){
        res.push_back(a[r][c]); used[r][c]=1;
        int nr=r+dr[dir], nc=c+dc[dir];
        if(nr<0||nr>=m||nc<0||nc>=n||used[nr][nc]){ dir=(dir+1)%4; nr=r+dr[dir]; nc=c+dc[dir]; }
        r=nr; c=nc;
    }
    return res;
}
vector<int> optimal_spiral_order(vector<vector<int>>& a){
    int top=0, left=0, down=a.size()-1, right=a[0].size()-1;
    vector<int> res;
    while(top<=down && left<=right){
        for(int j=left;j<=right;++j) res.push_back(a[top][j]);
        top++;
        for(int i=top;i<=down;++i) res.push_back(a[i][right]);
        right--;
        if(top<=down){
            for(int j=right;j>=left;--j) res.push_back(a[down][j]);
            down--;
        }
        if(left<=right){
            for(int i=down;i>=top;--i) res.push_back(a[i][left]);
            left++;
        }
    }
    return res;
}
int main(){return 0;}