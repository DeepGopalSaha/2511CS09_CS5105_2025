#include <bits/stdc++.h>
using namespace std;
void brute_set_matrix_zeroes(vector<vector<int>>& mat){
    int n=mat.size(), m=mat[0].size();
    vector<vector<int>> copy=mat;
    for(int i=0;i<n;++i) for(int j=0;j<m;++j) if(copy[i][j]==0){
        for(int c=0;c<m;++c) mat[i][c]=0;
        for(int r=0;r<n;++r) mat[r][j]=0;
    }
}
void optimal_set_matrix_zeroes(vector<vector<int>>& mat){
    int n=mat.size(), m=mat[0].size();
    bool row0=false, col0=false;
    for(int j=0;j<m;++j) if(mat[0][j]==0) row0=true;
    for(int i=0;i<n;++i) if(mat[i][0]==0) col0=true;
    for(int i=1;i<n;++i) for(int j=1;j<m;++j) if(mat[i][j]==0){ mat[i][0]=0; mat[0][j]=0; }
    for(int i=1;i<n;++i) for(int j=1;j<m;++j) if(mat[i][0]==0 || mat[0][j]==0) mat[i][j]=0;
    if(row0) for(int j=0;j<m;++j) mat[0][j]=0;
    if(col0) for(int i=0;i<n;++i) mat[i][0]=0;
}
int main(){return 0;}