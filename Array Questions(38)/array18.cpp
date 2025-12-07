#include <bits/stdc++.h>
using namespace std;
int brute_single_number(vector<int>& a){
    unordered_map<int,int> cnt;
    for(int x:a) cnt[x]++;
    for(auto &p:cnt) if(p.second==1) return p.first;
    return -1;
}
int optimal_single_number(vector<int>& a){
    int x=0; for(int v:a) x ^= v; return x;
}
int main(){return 0;}