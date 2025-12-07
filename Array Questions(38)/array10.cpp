#include <bits/stdc++.h>
using namespace std;
int brute_majority_element(vector<int>& a){
    int n=a.size();
    for(int i=0;i<n;++i){
        int cnt=0;
        for(int j=0;j<n;++j) if(a[j]==a[i]) cnt++;
        if(cnt>n/2) return a[i];
    }
    return -1;
}
int optimal_majority_element(vector<int>& a){
    int cand=0,c=0;
    for(int x:a){ if(c==0){cand=x; c=1;} else if(x==cand) c++; else c--; }
    return cand;
}
int main(){return 0;}