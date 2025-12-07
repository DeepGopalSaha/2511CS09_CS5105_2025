#include <bits/stdc++.h>
using namespace std;
int brute_remove_duplicates(vector<int>& a){
    vector<int> b=a;
    sort(b.begin(), b.end());
    b.erase(unique(b.begin(), b.end()), b.end());
    for(int i=0;i<b.size();++i) a[i]=b[i];
    return b.size();
}
int optimal_remove_duplicates(vector<int>& a){
    if(a.empty()) return 0;
    int i=0;
    for(int j=1;j<a.size();++j) if(a[j]!=a[i]) a[++i]=a[j];
    return i+1;
}
int main(){return 0;}