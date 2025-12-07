#include <bits/stdc++.h>
using namespace std;
void brute_move_zeros(vector<int>& a){
    vector<int> b;
    int zeros=0;
    for(int x:a) if(x==0) zeros++; else b.push_back(x);
    for(int i=0;i<zeros;i++) b.push_back(0);
    a=b;
}
void optimal_move_zeros(vector<int>& a){
    int pos=0;
    for(int i=0;i<a.size();++i) if(a[i]!=0) a[pos++]=a[i];
    while(pos<a.size()) a[pos++]=0;
}
int main(){return 0;}