#include <bits/stdc++.h>
using namespace std;
int brute_linear_search(vector<int>& a,int target){
    for(int i=0;i<a.size();++i) if(a[i]==target) return i;
    return -1;
}
int optimal_linear_search(vector<int>& a,int target){ return brute_linear_search(a,target); }
int main(){return 0;}